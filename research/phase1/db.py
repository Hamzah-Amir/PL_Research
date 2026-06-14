"""
PostgreSQL connection + schema for Phase 1.

Phase 1 is a stored-dataset model: Helium 10 Black Box exports are ingested into
Postgres, scored, and queried by budget (no live API at query time). This module
owns the connection (from `DATABASE_URL` in the project-root `.env`) and the
schema.

Schema = the agreed spec (products / product_history / product_reviews_insights)
plus a few fields needed by the rest of the tool:
  - `variation_count` / `has_variants`  — for the with/without-variants filter
    (Black Box "Variation Count": 0 = none, >0 = has variants).
  - `brand, fulfillment, image_url, url, listing_age_months, estimated_revenue,
    sales_trend_90d, price_trend_90d` — for the Excel shortlist + the Phase 2
    handoff (which needs the product's title/category) and richer scoring.

Raw SQL via psycopg2 (no ORM) — simple, fast, denormalised, per the project
guidance. No fallback: if the DB is unreachable, Phase 1 reports it.
"""

import os
import re
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
_ENV_PATH = _ROOT / ".env"


def _parse_db_url(url: str) -> dict:
    """Tolerantly parse a postgres URL into connect kwargs.

    Splits the host off the LAST '@', so a password containing unencoded '@'
    or ':' still parses (host/port/dbname never contain '@'). Avoids the need to
    percent-encode special characters in the password.
    """
    rest = re.sub(r"^\w+://", "", url.strip())
    if "@" not in rest:
        raise ValueError("DATABASE_URL missing user@host separator")
    userinfo, hostpart = rest.rsplit("@", 1)
    user, _, password = userinfo.partition(":")
    hostport, _, dbname = hostpart.partition("/")
    dbname = dbname.split("?")[0]
    host, _, port = hostport.partition(":")
    return {
        "host": host or "localhost",
        "port": port or "5432",
        "dbname": dbname,
        "user": user,
        "password": password,
    }


class Phase1DbError(RuntimeError):
    """Raised when the database is unavailable or a DB operation fails."""


SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS products (
    id                  SERIAL PRIMARY KEY,
    asin                TEXT UNIQUE NOT NULL,
    title               TEXT,
    brand               TEXT,
    category            TEXT,
    subcategory         TEXT,
    price               NUMERIC,
    bsr                 INTEGER,
    estimated_sales     NUMERIC,    -- ASIN-level monthly sales
    estimated_revenue   NUMERIC,    -- ASIN-level monthly revenue
    review_count        INTEGER,
    rating              NUMERIC,
    weight_kg           NUMERIC,
    yoy_growth          NUMERIC,
    is_seasonal         BOOLEAN,
    variation_count     INTEGER,
    has_variants        BOOLEAN,
    is_amazon_seller    BOOLEAN,
    is_compatibility    BOOLEAN,
    is_global_brand     BOOLEAN,
    fulfillment         TEXT,
    image_url           TEXT,
    url                 TEXT,
    listing_age_months  NUMERIC,
    sales_trend_90d     NUMERIC,
    price_trend_90d     NUMERIC,
    price_range         TEXT,       -- low / mid / high
    review_bucket       TEXT,       -- low / medium / high
    sales_bucket        TEXT,       -- low / medium / high
    competition_score   NUMERIC,
    demand_score        NUMERIC,
    opportunity_score   NUMERIC,
    llm_adj             NUMERIC,    -- Claude judgement adjustment (-10 to +10)
    llm_reason          TEXT,       -- Claude's one-line reason for adjustment
    llm_at              TIMESTAMP,  -- When Claude scoring was applied
    created_at          TIMESTAMP NOT NULL DEFAULT NOW(),
    last_updated        TIMESTAMP NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_products_category      ON products (category);
CREATE INDEX IF NOT EXISTS idx_products_price_range   ON products (price_range);
CREATE INDEX IF NOT EXISTS idx_products_opportunity   ON products (opportunity_score DESC);
CREATE INDEX IF NOT EXISTS idx_products_has_variants  ON products (has_variants);
CREATE INDEX IF NOT EXISTS idx_products_is_seasonal   ON products (is_seasonal);

CREATE TABLE IF NOT EXISTS product_history (
    id              SERIAL PRIMARY KEY,
    asin            TEXT NOT NULL,
    price           NUMERIC,
    bsr             INTEGER,
    estimated_sales NUMERIC,
    review_count    INTEGER,
    captured_at     TIMESTAMP NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_history_asin ON product_history (asin);

CREATE TABLE IF NOT EXISTS product_reviews_insights (
    id                SERIAL PRIMARY KEY,
    asin              TEXT NOT NULL,
    common_complaints TEXT,
    negative_keywords JSONB,
    created_at        TIMESTAMP NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_insights_asin ON product_reviews_insights (asin);

-- Columns added after the initial build (idempotent for existing tables).
ALTER TABLE products ADD COLUMN IF NOT EXISTS weight_kg        NUMERIC;
ALTER TABLE products ADD COLUMN IF NOT EXISTS yoy_growth       NUMERIC;
ALTER TABLE products ADD COLUMN IF NOT EXISTS is_amazon_seller BOOLEAN;
ALTER TABLE products ADD COLUMN IF NOT EXISTS is_compatibility BOOLEAN;
ALTER TABLE products ADD COLUMN IF NOT EXISTS is_global_brand  BOOLEAN;
ALTER TABLE products ADD COLUMN IF NOT EXISTS llm_adj          NUMERIC;
ALTER TABLE products ADD COLUMN IF NOT EXISTS llm_reason       TEXT;
ALTER TABLE products ADD COLUMN IF NOT EXISTS llm_at           TIMESTAMP;
"""


def _connect_autocreate(psycopg2, kwargs: dict):
    """Connect; if the target database does not exist, create it (via the
    `postgres` maintenance DB) and retry. Makes first run seamless."""
    try:
        return psycopg2.connect(**kwargs)
    except psycopg2.OperationalError as e:
        dbname = kwargs.get("dbname")
        if not dbname or "does not exist" not in str(e):
            raise
        admin = dict(kwargs, dbname="postgres")
        conn = psycopg2.connect(**admin)
        try:
            conn.autocommit = True
            with conn.cursor() as cur:
                cur.execute(f'CREATE DATABASE "{dbname}"')
        finally:
            conn.close()
        print(f"  Created database '{dbname}'.")
        return psycopg2.connect(**kwargs)


def get_connection():
    """Open a psycopg2 connection from `.env` (DATABASE_URL or DB_* vars)."""
    try:
        from dotenv import load_dotenv
    except ImportError as e:
        raise Phase1DbError("python-dotenv is not installed. Run setup again.") from e
    try:
        import psycopg2
    except ImportError as e:
        raise Phase1DbError(
            "psycopg2 is not installed. Run setup again (pip install -r requirements.txt)."
        ) from e

    load_dotenv(dotenv_path=_ENV_PATH)

    # Preferred: discrete vars (no URL-encoding needed for special chars in the
    # password). Fall back to DATABASE_URL if discrete vars are not set.
    host = os.environ.get("DB_HOST")
    name = os.environ.get("DB_NAME")
    user = os.environ.get("DB_USER")
    url = os.environ.get("DATABASE_URL")

    if host and name and user:
        kwargs = {
            "host": host, "port": os.environ.get("DB_PORT", "5432"),
            "dbname": name, "user": user,
            "password": os.environ.get("DB_PASSWORD", ""),
        }
    elif url:
        # Parse tolerantly so special chars in the password (@ : /) work
        # without percent-encoding.
        kwargs = _parse_db_url(url)
    else:
        kwargs = None

    if kwargs is not None:
        try:
            return _connect_autocreate(psycopg2, kwargs)
        except Exception as e:  # noqa: BLE001
            raise Phase1DbError(
                f"Could not connect to PostgreSQL: {e}\n"
                f"If your password has special characters (@ : / etc.), either "
                f"percent-encode them in DATABASE_URL (@ -> %40) OR use the discrete "
                f"vars DB_HOST / DB_PORT / DB_NAME / DB_USER / DB_PASSWORD in {_ENV_PATH}."
            ) from e

    raise Phase1DbError(
        f"No database config found. Add to {_ENV_PATH} either:\n"
        f"  DATABASE_URL=postgresql://user:password@localhost:5432/pl_research\n"
        f"(percent-encode special chars in the password), or the discrete vars:\n"
        f"  DB_HOST=localhost\n  DB_PORT=5432\n  DB_NAME=pl_research\n"
        f"  DB_USER=postgres\n  DB_PASSWORD=your_password\n(the file is git-ignored)."
    )


def init_db(conn) -> None:
    """Create tables + indexes if they do not exist (idempotent)."""
    try:
        with conn.cursor() as cur:
            cur.execute(SCHEMA_SQL)
        conn.commit()
    except Exception as e:  # noqa: BLE001
        conn.rollback()
        raise Phase1DbError(f"Schema initialisation failed: {e}") from e
