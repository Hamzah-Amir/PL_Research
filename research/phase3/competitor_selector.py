"""
Competitor identification + selection (PDF Step 8).

Deterministic data ops here; the same-product-type judgement is Claude's
(`claude_client.filter_same_product_type`). The split, per the project rule:

  - Python: merge JS + the per-keyword Xray exports into one record per ASIN,
    compute a keyword-presence relevancy proxy (how many of the launch-keyword
    searches the ASIN appears in) and resolve the sales figure, then rank.
  - Claude: decide which candidates are the SAME product type as the target.

Selection rule (Step 8): include BOTH FBA and FBM; keep only same-product-type
ASINs; select the top 3 by monthly sales (tie-break on keyword presence, then
review count). Output carries the FBA/FBM flag, the full merged data and a
reason per pick.
"""

from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

from research.phase3.claude_client import ProductProfile, filter_same_product_type


def _num(v) -> Optional[float]:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _norm_fulfillment(raw) -> str:
    """Normalise fulfilment codes to FBA / FBM / AMZ across JS and Xray."""
    s = str(raw or "").strip().upper()
    if s in ("FBA",):
        return "FBA"
    if s in ("FBM", "MFN"):
        return "FBM"
    if s in ("AMZ", "AMAZON"):
        return "AMZ"
    return s or "Unknown"


def build_candidate_records(
    js_df: Optional[pd.DataFrame],
    per_keyword_xray: Dict[str, pd.DataFrame],
    sponsored_agg: Dict[str, Dict],
    pool: List[str],
) -> pd.DataFrame:
    """One merged record per ASIN in *pool* (JS + Xray fields + sponsored agg).

    Sales source priority: JS monthly_sales (the workflow's estimate), else the
    Xray ASIN-level sales. For each ASIN the representative Xray row is the one
    with the highest ASIN sales across the keyword exports.
    """
    js_by_asin: Dict[str, pd.Series] = {}
    if js_df is not None and "asin" in js_df.columns:
        for _, r in js_df.iterrows():
            js_by_asin.setdefault(r["asin"], r)

    # Best Xray row per ASIN + presence across keyword files. Presence counts
    # DISTINCT keywords (an ASIN can appear as both an organic and a sponsored
    # row within one export, which must not double-count).
    xray_best: Dict[str, pd.Series] = {}
    presence_kw: Dict[str, set] = {}
    for kw, df in per_keyword_xray.items():
        if "asin" not in df.columns:
            continue
        for _, r in df.iterrows():
            a = r["asin"]
            presence_kw.setdefault(a, set()).add(kw)
            cur = xray_best.get(a)
            if cur is None or (_num(r.get("asin_sales")) or 0) > (_num(cur.get("asin_sales")) or 0):
                xray_best[a] = r
    presence = {a: len(kws) for a, kws in presence_kw.items()}

    records = []
    for asin in pool:
        js = js_by_asin.get(asin)
        xr = xray_best.get(asin)
        spons = sponsored_agg.get(asin, {})

        def pick(js_key, xr_key=None, prefer="js"):
            xr_key = xr_key or js_key
            jv = js.get(js_key) if js is not None else None
            xv = xr.get(xr_key) if xr is not None else None
            order = (jv, xv) if prefer == "js" else (xv, jv)
            for v in order:
                if v is not None and not (isinstance(v, float) and np.isnan(v)):
                    return v
            return None

        monthly_sales = _num(pick("monthly_sales", "asin_sales"))
        monthly_revenue = _num(pick("monthly_revenue", "asin_revenue"))
        fulfillment = _norm_fulfillment(pick("fulfillment", "fulfillment"))

        records.append({
            "asin": asin,
            "title": pick("title", "title"),
            "brand": pick("brand", "brand"),
            "price": _num(pick("price", "price")),
            "monthly_sales": monthly_sales,
            "monthly_revenue": monthly_revenue,
            "rating": _num(pick("rating", "rating", prefer="xr")),
            "review_count": _num(pick("reviews", "review_count", prefer="xr")),
            "review_velocity": _num(xr.get("review_velocity")) if xr is not None else None,
            "best_seller": (str(xr.get("best_seller")).strip() if xr is not None and xr.get("best_seller") is not None else None),
            "bsr": _num(pick("bsr", "bsr", prefer="xr")),
            "date_first_available": pick("date_first_available", "creation_date"),
            "fulfillment": fulfillment,
            "seller": xr.get("seller") if xr is not None else None,
            "seller_age_mo": _num(xr.get("seller_age_mo")) if xr is not None else None,
            "active_sellers": _num(pick("num_sellers", "active_sellers", prefer="xr")),
            "image_url": (xr.get("image_url") if xr is not None else None),
            "url": pick("link", "url", prefer="xr"),
            "dimensions": pick("dimensions", "dimensions"),
            "weight": pick("weight", "weight"),
            "category": pick("category", "category"),
            "sp_count": spons.get("sp_count", 0),
            "sb_count": spons.get("sb_count", 0),
            "sp_keywords": spons.get("sp_keywords", []),
            "sb_keywords": spons.get("sb_keywords", []),
            "keyword_presence": presence.get(asin, 0),
            "total_keywords": len(per_keyword_xray),
            "in_js": js is not None,
            "in_xray": xr is not None,
        })

    return pd.DataFrame(records)


def select_top_competitors(
    client,
    profile: ProductProfile,
    records: pd.DataFrame,
    n: int = 3,
) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    """
    Return (top_n, all_same_type, report_lines).

    Claude filters to same-product-type; Python ranks the survivors by monthly
    sales (tie-break: keyword presence, then review count) and takes the top *n*.
    Both FBA and FBM are eligible — fulfilment is never used to exclude.
    """
    report: List[str] = []
    if records.empty:
        raise ValueError("No candidate records to select from.")

    result = filter_same_product_type(client, profile, records[["asin", "title"]])
    kept = set(result.same_type_asins)
    same = records[records["asin"].isin(kept)].copy()
    report.append(f"  Candidates considered : {len(records)}")
    report.append(f"  Same product type     : {len(same)}  (Claude)")
    report.append(f"  Note                  : {result.note}")

    if same.empty:
        return same, same, report

    same["_sales"] = same["monthly_sales"].fillna(0)
    same["_pres"] = same["keyword_presence"].fillna(0)
    same["_rev"] = same["review_count"].fillna(0)
    same = same.sort_values(["_sales", "_pres", "_rev"], ascending=False).reset_index(drop=True)
    same = same.drop(columns=["_sales", "_pres", "_rev"])

    top = same.head(n).copy()
    ff = ", ".join(f"{r.asin}={r.fulfillment}" for r in top.itertuples())
    report.append(f"  Top {len(top)} by sales      : {ff}")
    return top, same, report
