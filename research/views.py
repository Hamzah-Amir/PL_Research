from pathlib import Path

from django.conf import settings
from django.http import FileResponse, Http404
from django.shortcuts import render

from products.models import Product
from .services.phase1 import run_query
from .services.phase2 import run_phase2
from .services.phase3 import ELEMENTS, build_js_sheet, run_phase3

PER_PAGE = 10


def research_dashboard(request):

    categories = Product.objects.values_list('category', flat=True).distinct()
    return render(request, 'research/dashboard.html', {
        'categories': categories,
    })


def shortlist(request):
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

    if product and request.method == 'POST':
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

        if not xray_path:
            ctx['error'] = 'Please upload your Helium 10 Xray keyword export to continue.'
        else:
            exclude = request.POST.getlist('exclude') or None
            result = run_phase2(asin, xray_path, exclude=exclude)
            ctx['result'] = result
            # Carry the main search term + accepted keywords to Phase 3 (the
            # JS-sheet thread can't cross requests, so we keep the term string).
            if result and result.get('selection'):
                request.session[f'phase3_{asin}'] = {
                    'term': result.get('main_search_term'),
                    'keywords': [k['keyword'] for k in result['selection']['keywords']],
                }

    return render(request, 'research/select.html', ctx)


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

        # Generate the Jungle Scout sheet from the Phase 2 main search term.
        try:
            jsdir = Path(settings.MEDIA_ROOT) / 'phase3_js'
            jsdir.mkdir(parents=True, exist_ok=True)
            js = build_js_sheet(keyword=sess['term'], out_path=str(jsdir / f'{asin}.csv'))
            js_path = js.get('out_path')
        except Exception as e:  # noqa: BLE001
            ctx['error'] = f'Could not generate the Jungle Scout sheet for "{sess["term"]}": {e}'
            return render(request, 'research/phase3.html', ctx)

        result = run_phase3(asin, js_file=js_path, xray_files=saved,
                            locked_keywords=sess.get('keywords'))
        ctx['result'] = result

        if result.get('scored'):
            recs = result.get('records', {})
            comps = []
            for s in result['scored']:
                rec = recs.get(s['asin'], {})
                rows = [{'label': label, 'score': s['elements'][key]['score'], 'max': mx}
                        for key, label, mx, _ in ELEMENTS if key in s.get('elements', {})]
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


def download_workbook(request, name):
    """Serve a generated Competitor Analysis deliverable (.xlsx / .pdf) from output/."""
    path = Path(settings.BASE_DIR) / 'output' / name
    if path.suffix not in ('.xlsx', '.pdf') or not path.exists():
        raise Http404('Deliverable not found.')
    return FileResponse(open(path, 'rb'), as_attachment=True, filename=name)