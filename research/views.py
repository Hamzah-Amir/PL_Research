from pathlib import Path

from django.conf import settings
from django.http import FileResponse, Http404
from django.shortcuts import redirect, render

from products.models import Product
from .services.phase1 import run_query
from .services.phase2 import run_phase2
from .services.phase3 import ELEMENTS, build_js_sheet, run_phase3
from .services.phase4 import COLUMNS as PHASE4_COLUMNS, run_phase4

PER_PAGE = 10


def research_dashboard(request):
    ctx = {'ingest': None}
    # Optional: add products to the DB from a Black Box Excel upload.
    if request.method == 'POST' and request.FILES.get('blackbox_file'):
        from .services.ingest import ingest_products_excel
        ctx['ingest'] = ingest_products_excel(request.FILES['blackbox_file'])
    ctx['categories'] = Product.objects.values_list('category', flat=True).distinct()
    return render(request, 'research/dashboard.html', ctx)


def shortlist(request):
    # Optional direct-ASIN shortcut: if the user supplied an ASIN on the
    # dashboard, skip the shortlist and jump straight to Phase 2 for it.
    asin = (request.GET.get('asin') or '').strip().upper()
    if asin:
        return redirect('research:select', asin=asin)

    budget = request.GET.get('budget', '')
    category = request.GET.get('category', '')
    include_seasonal = request.GET.get('include_seasonal') == '1'

    # Fetch the whole ranked pool once; the template renders 10 at a time and
    # JS pages through the rest with no further server round-trips.
    result = run_query(budget, category, include_seasonal, limit=0)
    products = result['products']

    return render(request, 'research/shortlist.html', {
        'budget': budget,
        'category': category,
        'include_seasonal': include_seasonal,
        'products': products,
        'pool_size': result['pool_size'],
        'per_page': PER_PAGE,
    })

def select_asin(request, asin):
    """Phase 2 for the chosen ASIN: upload the H10 Xray keyword export, then show
    the Claude product profile + top keywords. Re-running with rejected keywords
    (the swap form) re-uses the saved upload, so no re-upload is needed."""
    asin = (asin or '').strip().upper()
    product = Product.objects.filter(asin=asin).first()
    ctx = {'asin': asin, 'product': product, 'result': None, 'error': None}

    # Process the upload even when the ASIN isn't in the DB — run_phase2 falls
    # back to Keepa for the product details (dev use), so a directly-entered
    # ASIN still reaches keyword research.
    if request.method == 'POST':
        session_key = f'phase2_xray_{asin}'
        upload = request.FILES.get('xray_file')

        if upload:
            # Persist the upload so swap rounds don't require re-uploading.
            updir = Path(settings.MEDIA_ROOT) / 'phase2_xray'
            updir.mkdir(parents=True, exist_ok=True)
            saved = updir / f'{asin}.xlsx'
            with open(saved, 'wb') as fh:
                for chunk in upload.chunks():
                    fh.write(chunk)
            request.session[session_key] = str(saved)
            xray_path = str(saved)
        else:
            xray_path = request.session.get(session_key)

        # Optional Keepa export — only used for the not-in-DB fallback (no-API
        # path); the normal DB flow ignores it.
        keepa_path = _save_keepa_upload(request, asin, 'phase2')

        if not xray_path:
            ctx['error'] = 'Please upload your Helium 10 Xray keyword export to continue.'
        else:
            exclude = request.POST.getlist('exclude') or None
            result = run_phase2(asin, xray_path, exclude=exclude, keepa_file=keepa_path)
            ctx['result'] = result
            # Carry the main search term + accepted keywords to Phase 3 (the
            # JS-sheet thread can't cross requests, so we keep the term string).
            if result and result.get('selection'):
                request.session[f'phase3_{asin}'] = {
                    'term': result.get('main_search_term'),
                    'keywords': [k['keyword'] for k in result['selection']['keywords']],
                }

    return render(request, 'research/select.html', ctx)


def _save_keepa_upload(request, asin, slot):
    """Persist an uploaded Keepa export (field name ``keepa_file``) under
    ``media/<slot>_keepa/<asin>`` and remember it in the session so it survives
    multi-step POSTs. Returns the path, or None if none is uploaded/stored. The
    live Keepa API is used only when no file is present and USE_KEEPA_API is on."""
    key = f'{slot}_keepa_{asin}'
    up = request.FILES.get('keepa_file')
    path = request.session.get(key)
    if up:
        kdir = Path(settings.MEDIA_ROOT) / f'{slot}_keepa'
        kdir.mkdir(parents=True, exist_ok=True)
        path = str(kdir / f'{asin}{Path(up.name).suffix or ".xlsx"}')
        with open(path, 'wb') as fh:
            for chunk in up.chunks():
                fh.write(chunk)
        request.session[key] = path
    if not (path and Path(path).exists()):
        path = None
        request.session.pop(key, None)
    return path


def phase3(request, asin):
    """Phase 3 for the accepted ASIN: upload the H10 Xray *Products* exports (one
    per launch keyword); the Jungle Scout sheet is generated from the main search
    term carried from Phase 2. Runs the competitor analysis and writes the
    workbook."""
    asin = (asin or '').strip().upper()
    product = Product.objects.filter(asin=asin).first()
    sess = request.session.get(f'phase3_{asin}', {})
    ctx = {
        'asin': asin, 'product': product, 'term': sess.get('term'),
        'keywords': sess.get('keywords', []), 'result': None, 'error': None,
        'competitors': None, 'workbook_name': None, 'pdf_name': None,
    }

    if product and request.method == 'POST':
        if not sess.get('term'):
            ctx['error'] = 'Start from Phase 2 first — the main search term is needed to generate the Jungle Scout sheet.'
            return render(request, 'research/phase3.html', ctx)

        uploads = request.FILES.getlist('xray_files')
        if not uploads:
            ctx['error'] = 'Upload at least one Helium 10 Xray Products export (one per launch keyword).'
            return render(request, 'research/phase3.html', ctx)

        # Keepa competitor data comes from an uploaded Keepa export (no-API path);
        # the live Keepa API is used only when USE_KEEPA_API is on.
        keepa_path = _save_keepa_upload(request, asin, 'phase3')
        if not keepa_path and not getattr(settings, 'USE_KEEPA_API', False):
            ctx['error'] = ('Upload a Keepa export too — the competitor analysis needs Keepa data. '
                            'See "How to get the Keepa file" below. (Developers can set '
                            'USE_KEEPA_API=1 to use the live Keepa API instead.)')
            return render(request, 'research/phase3.html', ctx)

        # Save the Xray Products uploads.
        xdir = Path(settings.MEDIA_ROOT) / 'phase3_xray' / asin
        xdir.mkdir(parents=True, exist_ok=True)
        saved = []
        for f in uploads:
            dest = xdir / f.name
            with open(dest, 'wb') as fh:
                for chunk in f.chunks():
                    fh.write(chunk)
            saved.append(str(dest))

        # Generate the JS sheet from the Phase 2 term; fall back to a manual JS
        # sheet if scraping fails (Amazon block / no working proxy).
        try:
            js_path, js_note = _ensure_js_sheet(sess['term'], asin)
        except RuntimeError as e:
            ctx['error'] = f'Could not generate the Jungle Scout sheet for "{sess["term"]}": {e}'
            return render(request, 'research/phase3.html', ctx)

        result = run_phase3(asin, js_file=js_path, xray_files=saved,
                            locked_keywords=sess.get('keywords'), keepa_file=keepa_path)
        if js_note:
            result.setdefault('log', []).insert(0, js_note)
        ctx['result'] = result
        # Data-source failures (reviews API, Keepa, …) are buried in the run
        # log; surface them so a sheet with 'Unknown' cells isn't a mystery.
        ctx['warnings'] = [l for l in (result.get('log') or [])
                           if 'failed' in l.lower() or 'unavailable' in l.lower()]

        if result.get('scored'):
            recs = result.get('records', {})
            comps = []
            for s in result['scored']:
                rec = recs.get(s['asin'], {})
                rows = [{'label': label, 'score': s['elements'][key]['score'],
                         'max': s['elements'][key]['max']}
                        for key, label, _mx, _ in ELEMENTS if key in s.get('elements', {})]
                comps.append({
                    'asin': s['asin'], 'brand': rec.get('brand'), 'title': rec.get('title'),
                    'total': s['total'], 'max': s['max'], 'band': s['band'],
                    'unknown': s.get('unknown_elements'), 'rows': rows,
                })
            ctx['competitors'] = comps
        if result.get('workbook_path'):
            ctx['workbook_name'] = Path(result['workbook_path']).name
        if result.get('pdf_path'):
            ctx['pdf_name'] = Path(result['pdf_path']).name

    return render(request, 'research/phase3.html', ctx)


def _lines(raw):
    """Split a textarea value into a clean list (one item per non-empty line)."""
    return [ln.strip() for ln in (raw or '').splitlines() if ln.strip()]


def _ensure_js_sheet(term, asin):
    """Build the Jungle Scout sheet from *term*; if scraping fails (Amazon block
    / no working proxy) fall back to the manual JS sheet in
    ``settings.PHASE3_FALLBACK_JS``. Returns ``(js_path, note)`` — *note* is None
    on success, or a 'used fallback' message. Raises RuntimeError if neither the
    scrape nor a fallback file is available."""
    jsdir = Path(settings.MEDIA_ROOT) / 'phase3_js'
    jsdir.mkdir(parents=True, exist_ok=True)
    try:
        js = build_js_sheet(keyword=term, out_path=str(jsdir / f'{asin}.csv'))
        return js.get('out_path'), None
    except Exception as e:  # noqa: BLE001
        fallback = getattr(settings, 'PHASE3_FALLBACK_JS', '') or str(
            Path(settings.BASE_DIR) / 'test_file' / 'js-keyword-search.csv')
        if Path(fallback).exists():
            note = (f'JS-sheet scraping failed ({e}). Using the fallback JS sheet '
                    f'"{Path(fallback).name}" instead — set up a working proxy to '
                    f'auto-generate it, or replace the fallback file.')
            return fallback, note
        raise RuntimeError(f'{e} (and no fallback JS sheet at {fallback})')


def phase4(request, asin):
    """Phase 4 (Critical Sheet). Two-stage flow adapted from the v9.6 workflow:
    (1) Claude proposes a controlled vocabulary + Design attributes and the tool
    pulls the variation universe from Keepa; (2) after the user reviews/edits and
    approves the vocabulary, Claude classifies each parent and the Critical Sheet
    workbook is written. The Jungle Scout sheet is regenerated (cached) from the
    Phase 2 main search term carried in the session."""
    asin = (asin or '').strip().upper()
    product = Product.objects.filter(asin=asin).first()
    sess = request.session.get(f'phase3_{asin}', {})
    ctx = {
        'asin': asin, 'product': product, 'term': sess.get('term'),
        'result': None, 'error': None, 'columns': [h for _k, h in PHASE4_COLUMNS],
    }

    if not (product and request.method == 'POST'):
        return render(request, 'research/phase4.html', ctx)

    if not sess.get('term'):
        ctx['error'] = 'Start from Phase 2/3 first — the main search term is needed to build the Jungle Scout universe.'
        return render(request, 'research/phase4.html', ctx)

    action = request.POST.get('action', 'propose')

    # ── Phase 6: fill the KWs tabs into the already-built workbook. Runs AFTER
    # the Critical Sheet (the top-10 ASINs come from it), so no JS/Keepa/Claude. ─
    if action == 'cerebro':
        from .services.phase4 import add_kws_to_workbook
        wb = request.session.get(f'phase4_workbook_{asin}')
        up = request.FILES.get('cerebro_file')
        if not wb or not Path(wb).exists():
            ctx['error'] = 'Build the Critical Sheet first, then upload the Cerebro export.'
        elif not up:
            ctx['error'] = 'Please choose a Cerebro export (.xlsx/.csv) to upload.'
        else:
            cdir = Path(settings.MEDIA_ROOT) / 'phase4_cerebro'
            cdir.mkdir(parents=True, exist_ok=True)
            cpath = cdir / f'{asin}{Path(up.name).suffix or ".xlsx"}'
            with open(cpath, 'wb') as fh:
                for chunk in up.chunks():
                    fh.write(chunk)
            log = []
            try:
                add_kws_to_workbook(wb, str(cpath), asin, sess.get('keywords'), log=log)
                ctx['result'] = {'stage': 'cerebro', 'workbook_path': wb, 'log': log,
                                 'top10_asins': request.session.get(f'phase4_top10_{asin}', [])}
                ctx['workbook_name'] = Path(wb).name
                ctx['cerebro_done'] = True
            except Exception as e:  # noqa: BLE001
                ctx['error'] = f'Could not fill the KWs tabs from the Cerebro export: {e}'
        return render(request, 'research/phase4.html', ctx)

    # ── Keepa source: no-API path. The user uploads a Keepa website export; we
    # persist it (propose + build are two separate POSTs) and reuse it. The live
    # Keepa API is only used when settings.PHASE4_USE_KEEPA_API is on. ──────────
    keepa_path = _save_keepa_upload(request, asin, 'phase4')
    if not keepa_path and not getattr(settings, 'USE_KEEPA_API', False):
        ctx['error'] = ('Upload a Keepa export first — the Critical Sheet needs Keepa product '
                        'data. See "How to get the Keepa file" below. (Developers can set '
                        'USE_KEEPA_API=1 to use the live Keepa API instead.)')
        return render(request, 'research/phase4.html', ctx)

    # The JS sheet is the same one Phase 3 uses; regenerate it (cached → free),
    # falling back to the manual JS sheet if scraping fails (no working proxy).
    try:
        js_path, js_note = _ensure_js_sheet(sess['term'], asin)
    except RuntimeError as e:
        ctx['error'] = f'Could not generate the Jungle Scout sheet for "{sess["term"]}": {e}'
        return render(request, 'research/phase4.html', ctx)
    approved = None
    if action == 'build':
        approved = {
            'material_values': _lines(request.POST.get('material_values')),
            'size_guidance': (request.POST.get('size_guidance') or '').strip(),
            'color_values': _lines(request.POST.get('color_values')),
            'packaging_values': _lines(request.POST.get('packaging_values')),
            'special_feature_guidance': (request.POST.get('special_feature_guidance') or '').strip(),
            'design_attributes': _lines(request.POST.get('design_attributes')),
        }

    # Xray exports uploaded in Phase 3 feed the H10 Basic Data tab (top 3 keywords).
    xray_dir = Path(settings.MEDIA_ROOT) / 'phase3_xray' / asin
    xray_files = str(xray_dir) if xray_dir.exists() else None

    # Build the Critical Sheet (Phase 4/5). Cerebro/KWs (Phase 6) is a SEPARATE
    # step afterwards — the top-10 ASINs it needs only exist once this is built.
    result = run_phase4(asin, js_file=js_path, keyword=sess['term'], approved=approved,
                        xray_files=xray_files, launch_keywords=sess.get('keywords'),
                        keepa_file=keepa_path)
    if js_note:
        result.setdefault('log', []).insert(0, js_note)
    ctx['result'] = result
    ctx['warnings'] = [l for l in (result.get('log') or [])
                       if 'failed' in l.lower() or 'unavailable' in l.lower()]
    if result.get('workbook_path'):
        ctx['workbook_name'] = Path(result['workbook_path']).name
        # Persist for the Phase-6 Cerebro step (run Cerebro on these top-10 ASINs).
        request.session[f'phase4_workbook_{asin}'] = result['workbook_path']
        request.session[f'phase4_top10_{asin}'] = result.get('top10_asins', [])
    return render(request, 'research/phase4.html', ctx)


def download_workbook(request, name):
    """Serve a generated Competitor Analysis deliverable (.xlsx / .pdf) from output/."""
    path = Path(settings.BASE_DIR) / 'output' / name
    if path.suffix not in ('.xlsx', '.pdf') or not path.exists():
        raise Http404('Deliverable not found.')
    return FileResponse(open(path, 'rb'), as_attachment=True, filename=name)