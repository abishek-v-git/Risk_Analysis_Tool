import re
import pandas as pd
from django.shortcuts import render
from django.contrib.auth.decorators import login_required
from django.core.files.storage import FileSystemStorage
from django.core.cache import cache
from django.http import JsonResponse
from django.conf import settings
import os
import datetime
import glob
from rapidfuzz.distance import Levenshtein
from .models import ReportAnalysis


def ensure_limit(reports_dir):
    all_hist = glob.glob(os.path.join(reports_dir, "report_*.xlsx"))
    all_hist.sort(key=os.path.getmtime, reverse=True)
    if len(all_hist) > 5:
        for old in all_hist[5:]:
            if os.path.isfile(old):
                os.remove(old)


def parse_excel(file_path):
    """Parse an Excel file and return a cleaned, column-reordered DataFrame.

    Column order matches what the front-end expects:
    Part Number → CPN → remaining columns (georisk removed).
    Reads the file only once.
    """
    df_raw = pd.read_excel(file_path, header=None)
    header_row_idx = None
    for idx, row in df_raw.iterrows():
        row_values = [str(cell).lower().strip() if pd.notna(cell) else '' for cell in row]
        key_columns = ['part number', 'manufacturer', 'cpn', 'georisk']
        matches = sum(1 for col in key_columns if any(col in val for val in row_values))
        if matches >= 3:
            header_row_idx = idx
            break

    h = header_row_idx if header_row_idx is not None else 0
    df = df_raw.iloc[h + 1:].copy()
    df.columns = df_raw.iloc[h].tolist()
    df = df.reset_index(drop=True)

    # Replace '?' in CPN column
    col_lower = [str(col).lower().strip() for col in df.columns]
    cpn_col_idx = next((i for i, c in enumerate(col_lower) if 'cpn' in c), None)
    if cpn_col_idx is not None:
        df.iloc[:, cpn_col_idx] = df.iloc[:, cpn_col_idx].replace('?', 'QMR123')

    df = df.dropna(how='all').loc[:, df.columns.notna()]
    df = df[[c for c in df.columns if not str(c).startswith('Unnamed')]].reset_index(drop=True).fillna('')

    # Reorder columns: remove georisk, move CPN to front, then Part Number to front
    cols = list(df.columns)
    col_lower = [str(c).lower().strip() for c in cols]

    if 'georisk' in col_lower:
        i = col_lower.index('georisk')
        cols.pop(i)
        col_lower.pop(i)

    cpn_names = ['cpn', 'internal part', 'customer part']
    cpn_idx = next((i for i, c in enumerate(col_lower) if c in cpn_names), -1)
    if cpn_idx == -1:
        cpn_idx = next((i for i, c in enumerate(col_lower) if any(n in c for n in cpn_names if len(n) >= 3)), -1)
    if cpn_idx != -1:
        cpn_col = cols.pop(cpn_idx)
        cols.insert(0, cpn_col)
        col_lower.pop(cpn_idx)
        col_lower.insert(0, cpn_col.lower().strip())

    pn_names = ['part number', 'apn', 'mpn']
    pn_idx = next((i for i, c in enumerate(col_lower) if c in pn_names), -1)
    if pn_idx == -1:
        pn_idx = next((i for i, c in enumerate(col_lower) if any(n in c for n in pn_names if len(n) >= 3)), -1)
    if pn_idx != -1:
        pn_col = cols.pop(pn_idx)
        cols.insert(0, pn_col)

    return df[cols]


# ── Fuzzy matching helpers (mirror JS logic exactly) ──────────────────────────

_NOISE = re.compile(
    r'\(obs\)|\(strip\)|\(lf\)\(sn\)|lead free|leadfree|pbfree|pbf|nopb|tr'
)
_NON_ALNUM = re.compile(r'[^a-z0-9]')
_LEADING_ZEROS = re.compile(r'0+([1-9])')


def super_normalize(val):
    if not val:
        return ''
    s = str(val).lower()
    s = _NOISE.sub('', s)
    s = _NON_ALNUM.sub('', s)
    s = _LEADING_ZEROS.sub(r'\1', s)
    return s.strip()


def _lev_ratio(n1, n2):
    dist = Levenshtein.distance(n1, n2)
    max_len = max(len(n1), len(n2))
    if max_len == 0:
        return 0.0
    return (max_len - dist) / max_len


def fuzzy_match(s1, s2, threshold=0.85):
    n1 = super_normalize(s1)
    n2 = super_normalize(s2)
    if not n1 or not n2:
        return False
    if n1 == n2:
        return True
    if len(n1) > 5 and len(n2) > 5 and (n1 in n2 or n2 in n1):
        return True
    return _lev_ratio(n1, n2) >= threshold


def similarity_score(s1, s2):
    n1 = super_normalize(s1)
    n2 = super_normalize(s2)
    if not n1 or not n2:
        return 0.0
    if n1 == n2:
        return 1.0
    if len(n1) > 5 and len(n2) > 5 and (n1 in n2 or n2 in n1):
        return 0.99
    return _lev_ratio(n1, n2)


def _make_step_spans(steps):
    parts = []
    for step in steps:
        lower = step.lower()
        if 'assembly' in lower:
            parts.append(f'<span class="assembly">{step}</span>')
        elif 'origin' in lower:
            parts.append(f'<span class="origin">{step}</span>')
        elif 'fabrication' in lower:
            parts.append(f'<span class="fabrication">{step}</span>')
        elif 'final' in lower:
            parts.append(f'<span class="final-test">{step}</span>')
        elif 'wafer' in lower:
            parts.append(f'<span class="wafer-test">{step}</span>')
        else:
            parts.append(step)
    return ', '.join(parts) if parts else 'no steps found'


def _is_risky(country):
    c = str(country).lower()
    return 'china' in c or 'taiwan' in c


# ── Views ─────────────────────────────────────────────────────────────────────

@login_required
def index(request):
    fs = FileSystemStorage()
    reports_dir = os.path.join(settings.MEDIA_ROOT, 'reports')
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)

    if request.GET.get('delete_file'):
        file_to_delete = request.GET.get('delete_file')
        ReportAnalysis.objects.filter(user=request.user, report_file_path=file_to_delete).delete()
        file_full_path = os.path.join(reports_dir, file_to_delete)
        if os.path.isfile(file_full_path):
            os.remove(file_full_path)
        from django.shortcuts import redirect
        return redirect('/')

    if request.GET.get('clear') == '1':
        ReportAnalysis.objects.filter(user=request.user).delete()
        files = glob.glob(os.path.join(reports_dir, "*"))
        for f in files:
            if os.path.isfile(f):
                os.remove(f)
        from django.shortcuts import redirect
        return redirect('/')

    ensure_limit(reports_dir)

    view_filename = request.GET.get('view_file')
    file_path = None
    original_name = None

    if view_filename:
        possible_path = os.path.join(reports_dir, view_filename)
        if os.path.isfile(possible_path):
            file_path = possible_path
            if view_filename.startswith("report_") and len(view_filename) > 22:
                original_name = view_filename[22:]
            else:
                original_name = view_filename

    try:
        is_new_upload = bool(request.method == 'POST' and request.FILES.get('excel_file'))
        if is_new_upload or file_path:
            if is_new_upload:
                file_obj = request.FILES['excel_file']
                original_name = file_obj.name

                fixed_filename = "uploaded_file.xlsx"
                if fs.exists(fixed_filename):
                    fs.delete(fixed_filename)
                fs.save(fixed_filename, file_obj)
                file_path = fs.path(fixed_filename)

                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                safe_name = "".join([c for c in original_name if c.isalnum() or c in ('.', '_')]).strip()
                hist_name = f"report_{timestamp}_{safe_name}"
                hist_path = os.path.join(reports_dir, hist_name)

                with open(hist_path, 'wb+') as dest:
                    for chunk in file_obj.chunks():
                        dest.write(chunk)

                ReportAnalysis.objects.create(
                    user=request.user,
                    filename=original_name,
                    report_file_path=hist_name
                )

                ensure_limit(reports_dir)
                file_path = hist_path

            try:
                df = parse_excel(file_path)
                columns = df.columns.tolist()

                # Cache the parsed DataFrame so the paginated API avoids re-parsing
                file_basename = os.path.basename(file_path)
                cache.set(f'table_{request.user.id}_{file_basename}', df, 3600)

                db_recent = ReportAnalysis.objects.filter(user=request.user)[:5]
                recent_list = []
                for entry in db_recent:
                    file_full_path = os.path.join(reports_dir, entry.report_file_path)
                    file_exists = os.path.isfile(file_full_path)
                    recent_list.append({
                        'name': entry.filename,
                        'file': entry.report_file_path,
                        'date': entry.upload_date.strftime("%Y-%m-%d %H:%M:%S"),
                        'status': 'complete' if file_exists else 'missing'
                    })

                if request.method == 'POST' and request.FILES.get('excel_file'):
                    from django.shortcuts import redirect
                    current_hist_name = hist_name if 'hist_name' in locals() else os.path.basename(file_path)
                    return redirect(f'/?view_file={current_hist_name}&analyze=1')

                return render(request, 'flagrisk/index.html', {
                    'columns': columns,
                    'is_uploaded': True,
                    'filename': original_name,
                    'recent_reports': recent_list,
                    'current_file': view_filename or (hist_name if 'hist_name' in locals() else None)
                })
            except Exception as e:
                return render(request, 'flagrisk/index.html', {'error': f"Error parsing Excel: {str(e)}"})

        db_recent = ReportAnalysis.objects.filter(user=request.user)[:5]
        recent_list = []
        for entry in db_recent:
            file_full_path = os.path.join(reports_dir, entry.report_file_path)
            file_exists = os.path.isfile(file_full_path)
            recent_list.append({
                'name': entry.filename,
                'file': entry.report_file_path,
                'date': entry.upload_date.strftime("%Y-%m-%d %H:%M:%S"),
                'status': 'complete' if file_exists else 'missing'
            })

        return render(request, 'flagrisk/index.html', {'recent_reports': recent_list})

    except Exception as e:
        return render(request, 'flagrisk/index.html', {'error': f"General error: {str(e)}"})


@login_required
def table_data_api(request):
    """Server-side DataTables endpoint. Supports pagination, search, and ordering.

    Pass length=-1 to return all rows (used by the analyze function).
    """
    file_name = request.GET.get('file', '').strip()
    draw = int(request.GET.get('draw', 1))
    start = int(request.GET.get('start', 0))
    length = int(request.GET.get('length', 50))
    search_value = request.GET.get('search[value]', '').strip()
    order_col = int(request.GET.get('order[0][column]', 0))
    order_dir = request.GET.get('order[0][dir]', 'asc')

    if not file_name:
        return JsonResponse({'error': 'No file specified'}, status=400)

    cache_key = f'table_{request.user.id}_{file_name}'
    df = cache.get(cache_key)

    if df is None:
        reports_dir = os.path.join(settings.MEDIA_ROOT, 'reports')
        file_path = os.path.join(reports_dir, file_name)
        if not os.path.isfile(file_path):
            return JsonResponse({
                'draw': draw, 'recordsTotal': 0,
                'recordsFiltered': 0, 'data': [],
                'error': 'File not found'
            })
        try:
            df = parse_excel(file_path)
            cache.set(cache_key, df, 3600)
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)

    total = len(df)

    if search_value:
        mask = df.apply(
            lambda row: row.astype(str).str.contains(search_value, case=False, na=False).any(),
            axis=1
        )
        filtered_df = df[mask]
    else:
        filtered_df = df

    filtered_count = len(filtered_df)

    if 0 <= order_col < len(df.columns):
        col_name = df.columns[order_col]
        filtered_df = filtered_df.sort_values(
            by=col_name,
            ascending=(order_dir == 'asc'),
            key=lambda s: s.astype(str).str.lower()
        )

    page_df = filtered_df if length < 0 else filtered_df.iloc[start:start + length]

    data = page_df.astype(str).replace({'nan': '', 'None': ''}).values.tolist()

    return JsonResponse({
        'draw': draw,
        'recordsTotal': total,
        'recordsFiltered': filtered_count,
        'data': data,
    })


@login_required
def analyze_api(request):
    """Run fuzzy matching and risk classification server-side.

    Returns pre-computed newData, risk result tables, dashboard counts, and
    chip counts so the browser only has to render — no heavy JS loops needed.
    """
    file_name = request.GET.get('file', '').strip()
    if not file_name:
        return JsonResponse({'error': 'No file specified'}, status=400)

    cache_key = f'table_{request.user.id}_{file_name}'
    df = cache.get(cache_key)

    if df is None:
        reports_dir = os.path.join(settings.MEDIA_ROOT, 'reports')
        file_path = os.path.join(reports_dir, file_name)
        if not os.path.isfile(file_path):
            return JsonResponse({'error': 'File not found'}, status=404)
        try:
            df = parse_excel(file_path)
            cache.set(cache_key, df, 3600)
        except Exception as e:
            return JsonResponse({'error': str(e)}, status=500)

    headers = [str(c).lower().strip() for c in df.columns]

    def find_col(names, partial=False):
        for name in names:
            for i, h in enumerate(headers):
                if partial and name in h:
                    return i
                elif not partial and h == name:
                    return i
        return -1

    idx_pn = find_col(['part number'])
    idx_mfr = find_col(['manufacturer'])
    idx_cpn = find_col(['cpn', 'internal part', 'customer part'], partial=True)
    idx_country = find_col(['country'])
    idx_up_part = find_col(['uploaded part'])
    idx_up_mfg = find_col(['uploaded mfg'])
    idx_mfg_step = find_col(['manufacturing step'])
    idx_mpn_risk = find_col(['mpn risk'])

    if idx_pn == -1 or idx_mfr == -1 or idx_up_part == -1 or idx_up_mfg == -1:
        return JsonResponse({'error': 'Required columns (Part Number, Manufacturer, Uploaded Part, Uploaded Mfg) not found'}, status=400)

    # Columns before any previously appended risk columns
    slice_idx = idx_mpn_risk if idx_mpn_risk > -1 else len(df.columns)
    original_columns = list(df.columns[:slice_idx])

    all_data = df.astype(str).replace({'nan': '', 'None': ''}).values.tolist()

    # ── Phase 1: fuzzy matching ────────────────────────────────────────────────
    def upload_key(up_part, up_mfg):
        p = str(up_part or '').strip()
        m = str(up_mfg or '').strip()
        if not p and not m:
            return ''
        return (p + '||' + m).lower()

    upload_summary = {}
    matched_rows = []

    for row in all_data:
        sys_pn = str(row[idx_pn]).strip() if idx_pn < len(row) else ''
        up_pn = str(row[idx_up_part]).strip() if idx_up_part < len(row) else ''
        sys_mfr = str(row[idx_mfr]).strip() if idx_mfr < len(row) else ''
        up_mfr = str(row[idx_up_mfg]).strip() if idx_up_mfg < len(row) else ''

        key = upload_key(up_pn, up_mfr)
        if key and key not in upload_summary:
            upload_summary[key] = {
                'upPart': up_pn, 'upMfg': up_mfr,
                'matched': False, 'bestPn': '', 'bestMfr': '', 'bestScore': -1
            }

        if key:
            score = (similarity_score(sys_pn, up_pn) + similarity_score(sys_mfr, up_mfr)) / 2
            u = upload_summary[key]
            if score > u['bestScore']:
                u['bestScore'] = score
                u['bestPn'] = str(row[idx_pn]).strip()
                u['bestMfr'] = str(row[idx_mfr]).strip()

        is_pn_match = fuzzy_match(sys_pn, up_pn)
        is_mfr_match = fuzzy_match(sys_mfr, up_mfr)

        if sys_pn and sys_mfr and up_pn and up_mfr and is_pn_match and is_mfr_match:
            matched_rows.append(row)
            if key:
                upload_summary[key]['matched'] = True

    # Unmatched uploads (only partial mismatches — skip rows where BOTH are wrong)
    unmatched_uploads = []
    for u in upload_summary.values():
        if u['matched']:
            continue
        labels = []
        if not fuzzy_match(u['bestPn'], u['upPart']):
            labels.append('Part Number Mismatch')
        if not fuzzy_match(u['bestMfr'], u['upMfg']):
            labels.append('Manufacturer Mismatch')
        if len(labels) == 2:
            continue
        unmatched_uploads.append([u['bestPn'], u['upPart'], u['bestMfr'], u['upMfg'], ' & '.join(labels)])

    if not matched_rows:
        return JsonResponse({
            'newData': [], 'columns': original_columns,
            'highRiskResults': [], 'noRiskResults': [], 'lowRiskResults': [],
            'unmatchedUploads': unmatched_uploads, 'cpnRiskMap': {},
            'dashboard': {'cpn_high': 0, 'cpn_low': 0, 'cpn_no': 0, 'mpn_risk_map': {}},
            'chip_counts': {'high': 0, 'low': 0, 'none': 0},
        })

    # ── Phase 2: risk classification ──────────────────────────────────────────
    mpn_map = {}       # pn_norm  -> [countries]
    cpn_to_mpns = {}   # cpn      -> set of pn_norm
    mpn_site_map = {}  # pn_norm  -> [mfg_steps for risky countries]

    for row in matched_rows:
        pn_norm = super_normalize(str(row[idx_pn]).strip())
        cpn = str(row[idx_cpn]).strip().lower() if idx_cpn > -1 and idx_cpn < len(row) else ''
        country = str(row[idx_country]).strip().lower() if idx_country > -1 and idx_country < len(row) else ''
        mfg_step = str(row[idx_mfg_step]).strip() if idx_mfg_step > -1 and idx_mfg_step < len(row) else ''

        if pn_norm:
            mpn_map.setdefault(pn_norm, []).append(country)
            if _is_risky(country):
                steps = mpn_site_map.setdefault(pn_norm, [])
                if mfg_step and mfg_step not in steps:
                    steps.append(mfg_step)

        if cpn and pn_norm:
            cpn_to_mpns.setdefault(cpn, set()).add(pn_norm)

    mpn_risk_map = {
        pn: ('High Risk' if all(_is_risky(c) for c in countries) else 'No Risk')
        for pn, countries in mpn_map.items()
    }

    cpn_risk_map = {}
    for cpn, mpns in cpn_to_mpns.items():
        risks = [mpn_risk_map.get(m, 'No Risk') for m in mpns]
        if all(r == 'No Risk' for r in risks):
            cpn_risk_map[cpn] = 'No Risk'
        elif all(r == 'High Risk' for r in risks):
            cpn_risk_map[cpn] = 'High Risk'
        elif any(r == 'High Risk' for r in risks):
            cpn_risk_map[cpn] = 'Low Risk'
        else:
            cpn_risk_map[cpn] = 'No Risk'

    # ── Phase 3: build newData and risk result lists ───────────────────────────
    new_data = []
    high_risk_results, high_risk_seen = [], set()
    no_risk_results, no_risk_seen = [], set()
    low_risk_results, low_risk_seen = [], set()

    for row in all_data:
        sys_pn = str(row[idx_pn]).strip()
        up_pn = str(row[idx_up_part]).strip()
        sys_mfr = str(row[idx_mfr]).strip()
        up_mfr = str(row[idx_up_mfg]).strip()

        if not (sys_pn and sys_mfr and up_pn and up_mfr
                and fuzzy_match(sys_pn, up_pn)
                and fuzzy_match(sys_mfr, up_mfr)):
            continue

        pn_key = super_normalize(sys_pn)
        cpn = str(row[idx_cpn]).strip().lower() if idx_cpn > -1 and idx_cpn < len(row) else ''
        country = str(row[idx_country]).strip().lower() if idx_country > -1 and idx_country < len(row) else ''

        mpn_risk = mpn_risk_map.get(pn_key, 'No Risk')
        cpn_risk = cpn_risk_map.get(cpn, 'No Risk')

        key_pair = mpn_risk + '|' + cpn_risk
        if key_pair == 'No Risk|No Risk':
            remark = 'The CPN is completely safe from all risks.'
        elif key_pair == 'No Risk|Low Risk':
            remark = 'The CPN has non preferred countries but the MPN has atleast one preferred country, so low risk'
        elif key_pair == 'High Risk|Low Risk':
            steps_html = _make_step_spans(mpn_site_map.get(pn_key, []))
            remark = f'The CPN has non preferred country for {sys_pn} ({steps_html})'
        elif key_pair == 'High Risk|High Risk':
            steps_html = _make_step_spans(mpn_site_map.get(pn_key, []))
            remark = f'The CPN has non preferred country for {sys_pn} in all sites ({steps_html})'
        else:
            remark = f'CPN {cpn_risk.lower()} | MPN {mpn_risk.lower()}'

        base_row = list(row[:slice_idx])
        new_data.append(base_row + [remark, mpn_risk, cpn_risk])

        split_row = [
            row[idx_pn],
            row[idx_cpn] if idx_cpn > -1 and idx_cpn < len(row) else '',
            row[idx_country] if idx_country > -1 and idx_country < len(row) else '',
        ]
        risk_key = f"{split_row[0]}|{split_row[1]}|{split_row[2]}"

        if cpn_risk == 'High Risk':
            if risk_key not in high_risk_seen:
                high_risk_seen.add(risk_key)
                high_risk_results.append(split_row)
        elif cpn_risk == 'Low Risk':
            if risk_key not in low_risk_seen:
                low_risk_seen.add(risk_key)
                low_risk_results.append(split_row)
        else:
            if not _is_risky(country) and risk_key not in no_risk_seen:
                no_risk_seen.add(risk_key)
                no_risk_results.append(split_row)

    # ── Phase 4: dashboard counts ──────────────────────────────────────────────
    cpn_high = sum(1 for r in cpn_risk_map.values() if r == 'High Risk')
    cpn_low = sum(1 for r in cpn_risk_map.values() if r == 'Low Risk')
    cpn_no = sum(1 for r in cpn_risk_map.values() if r == 'No Risk')

    # MPN counts: de-duplicate by normalized PN, High Risk takes priority
    unique_high_mpns = set()
    unique_no_mpns = set()
    mpn_risk_col = slice_idx + 1  # position of MPN Risk in new_data rows

    for row in new_data:
        pn_key_d = super_normalize(str(row[idx_pn]).strip())
        if not pn_key_d:
            continue
        mpn_risk_val = row[mpn_risk_col]
        if mpn_risk_val == 'High Risk':
            unique_high_mpns.add(pn_key_d)
            unique_no_mpns.discard(pn_key_d)
        elif mpn_risk_val == 'No Risk' and pn_key_d not in unique_high_mpns:
            unique_no_mpns.add(pn_key_d)

    mpn_dashboard_risk_map = {k: 'High Risk' for k in unique_high_mpns}
    mpn_dashboard_risk_map.update({k: 'No Risk' for k in unique_no_mpns if k not in mpn_dashboard_risk_map})

    chip_high = sum(1 for r in new_data if r[-1] == 'High Risk')
    chip_low = sum(1 for r in new_data if r[-1] == 'Low Risk')
    chip_none = sum(1 for r in new_data if r[-1] == 'No Risk')

    return JsonResponse({
        'newData': new_data,
        'columns': original_columns,
        'highRiskResults': high_risk_results,
        'noRiskResults': no_risk_results,
        'lowRiskResults': low_risk_results,
        'unmatchedUploads': unmatched_uploads,
        'cpnRiskMap': cpn_risk_map,
        'dashboard': {
            'cpn_high': cpn_high,
            'cpn_low': cpn_low,
            'cpn_no': cpn_no,
            'mpn_risk_map': mpn_dashboard_risk_map,
        },
        'chip_counts': {'high': chip_high, 'low': chip_low, 'none': chip_none},
    })
