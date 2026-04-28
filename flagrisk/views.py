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
        user_reports = ReportAnalysis.objects.filter(user=request.user)
        for report in user_reports:
            file_full_path = os.path.join(reports_dir, report.report_file_path)
            if os.path.isfile(file_full_path):
                os.remove(file_full_path)
        user_reports.delete()
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
                        'status': 'Analyzed' if entry.is_analyzed else 'Pending'
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
                'status': 'Analyzed' if entry.is_analyzed else 'Pending'
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
    """Run optimized fuzzy matching and risk classification server-side."""
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

    # Pre-calculate column indices
    headers = [str(c).lower().strip() for c in df.columns]
    def find_col(names, partial=False):
        for name in names:
            for i, h in enumerate(headers):
                if (partial and name in h) or (not partial and h == name):
                    return i
        return -1

    cols = {
        'pn': find_col(['part number']),
        'mfr': find_col(['manufacturer']),
        'cpn': find_col(['cpn', 'internal part', 'customer part'], partial=True),
        'country': find_col(['country']),
        'up_part': find_col(['uploaded part']),
        'up_mfg': find_col(['uploaded mfg']),
        'mfg_step': find_col(['manufacturing step']),
        'mpn_risk': find_col(['mpn risk'])
    }

    if any(cols[k] == -1 for k in ['pn', 'mfr', 'up_part', 'up_mfg']):
        return JsonResponse({'error': 'Required columns missing (Part Number, Manufacturer, Uploaded Part/Mfg)'}, status=400)

    slice_idx = cols['mpn_risk'] if cols['mpn_risk'] > -1 else len(df.columns)
    original_columns = list(df.columns[:slice_idx])

    # --- Optimization: Pre-normalize columns in bulk ---
    norm_cache = {}
    def fast_norm(v):
        if v not in norm_cache:
            norm_cache[v] = super_normalize(v)
        return norm_cache[v]

    df_norm = pd.DataFrame()
    df_norm['pn'] = df.iloc[:, cols['pn']].astype(str).map(fast_norm)
    df_norm['up_pn'] = df.iloc[:, cols['up_part']].astype(str).map(fast_norm)
    df_norm['mfr'] = df.iloc[:, cols['mfr']].astype(str).map(fast_norm)
    df_norm['up_mfr'] = df.iloc[:, cols['up_mfg']].astype(str).map(fast_norm)

    # --- Optimized Matching ---
    all_data = df.astype(str).replace({'nan': '', 'None': ''}).values.tolist()
    upload_summary = {}
    matched_rows_indices = []

    for i in range(len(df)):
        sys_pn_n = df_norm.iat[i, 0]
        up_pn_n = df_norm.iat[i, 1]
        sys_mfr_n = df_norm.iat[i, 2]
        up_mfr_n = df_norm.iat[i, 3]

        if not up_pn_n and not up_mfr_n:
            continue

        ukey = f"{up_pn_n}||{up_mfr_n}"
        u = upload_summary.setdefault(ukey, {
            'upPart': all_data[i][cols['up_part']], 'upMfg': all_data[i][cols['up_mfg']],
            'matched': False, 'bestPn': '', 'bestMfr': '', 'bestScore': -1
        })

        # Score calculation with exact match short-circuit
        if sys_pn_n == up_pn_n and sys_mfr_n == up_mfr_n:
            score = 1.0
            is_match = True
        else:
            # Re-use normalized strings to avoid overhead
            s_pn = 1.0 if sys_pn_n == up_pn_n else (0.99 if (len(sys_pn_n) > 5 and (sys_pn_n in up_pn_n or up_pn_n in sys_pn_n)) else _lev_ratio(sys_pn_n, up_pn_n))
            s_mfr = 1.0 if sys_mfr_n == up_mfr_n else (0.99 if (len(sys_mfr_n) > 5 and (sys_mfr_n in up_mfr_n or up_mfr_n in sys_mfr_n)) else _lev_ratio(sys_mfr_n, up_mfr_n))
            score = (s_pn + s_mfr) / 2
            is_match = (s_pn >= 0.85 and s_mfr >= 0.85)

        if score > u['bestScore']:
            u['bestScore'] = score
            u['bestPn'] = all_data[i][cols['pn']]
            u['bestMfr'] = all_data[i][cols['mfr']]

        if is_match:
            u['matched'] = True
            matched_rows_indices.append(i)

    # --- Unmatched Uploads ---
    unmatched_uploads = []
    for u in upload_summary.values():
        if not u['matched']:
            labels = []
            if not fuzzy_match(u['bestPn'], u['upPart']): labels.append('Part Number Mismatch')
            if not fuzzy_match(u['bestMfr'], u['upMfg']): labels.append('Manufacturer Mismatch')
            if len(labels) < 2:
                unmatched_uploads.append([u['bestPn'], u['upPart'], u['bestMfr'], u['upMfg'], ' & '.join(labels)])

    if not matched_rows_indices:
        return JsonResponse({
            'newData': [], 'columns': original_columns,
            'highRiskResults': [], 'noRiskResults': [], 'lowRiskResults': [],
            'unmatchedUploads': unmatched_uploads, 'cpnRiskMap': {},
            'dashboard': {'cpn_high': 0, 'cpn_low': 0, 'cpn_no': 0, 'mpn_risk_map': {}},
            'chip_counts': {'high': 0, 'low': 0, 'none': 0},
        })

    # --- Risk Classification ---
    mpn_map = {}      # pn_norm -> [countries]
    cpn_to_mpns = {}  # cpn -> {pn_norm}
    mpn_site_map = {} # pn_norm -> [mfg_steps]

    for idx in matched_rows_indices:
        row = all_data[idx]
        pn_n = df_norm.iat[idx, 0]
        cpn = row[cols['cpn']].strip().lower() if cols['cpn'] > -1 else ''
        country = row[cols['country']].strip().lower() if cols['country'] > -1 else ''
        step = row[cols['mfg_step']].strip() if cols['mfg_step'] > -1 else ''

        if pn_n:
            mpn_map.setdefault(pn_n, []).append(country)
            if _is_risky(country):
                s_list = mpn_site_map.setdefault(pn_n, [])
                if step and step not in s_list: s_list.append(step)
            if cpn:
                cpn_to_mpns.setdefault(cpn, set()).add(pn_n)

    mpn_risk_map = {pn: ('High Risk' if all(_is_risky(c) for c in countries) else 'No Risk') for pn, countries in mpn_map.items()}
    cpn_risk_map = {}
    for cpn, mpns in cpn_to_mpns.items():
        risks = [mpn_risk_map.get(m, 'No Risk') for m in mpns]
        if all(r == 'No Risk' for r in risks): cpn_risk_map[cpn] = 'No Risk'
        elif all(r == 'High Risk' for r in risks): cpn_risk_map[cpn] = 'High Risk'
        elif any(r == 'High Risk' for r in risks): cpn_risk_map[cpn] = 'Low Risk'
        else: cpn_risk_map[cpn] = 'No Risk'

    # --- Build Results ---
    new_data = []
    hr_res, lr_res, nr_res = [], [], []
    seen_hr, seen_lr, seen_nr = set(), set(), set()

    for i in matched_rows_indices:
        row = all_data[i]
        pn_n = df_norm.iat[i, 0]
        cpn = row[cols['cpn']].strip().lower() if cols['cpn'] > -1 else ''
        country = row[cols['country']].strip().lower() if cols['country'] > -1 else ''

        m_risk = mpn_risk_map.get(pn_n, 'No Risk')
        c_risk = cpn_risk_map.get(cpn, 'No Risk')

        if m_risk == 'No Risk' and c_risk == 'No Risk': remark = 'The CPN is completely safe from all risks.'
        elif m_risk == 'No Risk' and c_risk == 'Low Risk': remark = 'The CPN has non preferred countries but the MPN has atleast one preferred country, so low risk'
        elif m_risk == 'High Risk':
            steps_h = _make_step_spans(mpn_site_map.get(pn_n, []))
            remark = f'The CPN has non preferred country for {row[cols["pn"]]} ({steps_h})'
            if c_risk == 'High Risk': remark += ' in all sites'
        else: remark = f'CPN {c_risk.lower()} | MPN {m_risk.lower()}'

        new_data.append(list(row[:slice_idx]) + [remark, m_risk, c_risk])

        r_key = f"{row[cols['pn']]}|{cpn}|{country}"
        split_row = [row[cols['pn']], row[cols['cpn']] if cols['cpn'] > -1 else '', row[cols['country']] if cols['country'] > -1 else '']
        
        if c_risk == 'High Risk' and r_key not in seen_hr:
            seen_hr.add(r_key); hr_res.append(split_row)
        elif c_risk == 'Low Risk' and r_key not in seen_lr:
            seen_lr.add(r_key); lr_res.append(split_row)
        elif c_risk == 'No Risk' and not _is_risky(country) and r_key not in seen_nr:
            seen_nr.add(r_key); nr_res.append(split_row)

    # Dashboard Counts
    dashboard = {
        'cpn_high': sum(1 for r in cpn_risk_map.values() if r == 'High Risk'),
        'cpn_low': sum(1 for r in cpn_risk_map.values() if r == 'Low Risk'),
        'cpn_no': sum(1 for r in cpn_risk_map.values() if r == 'No Risk'),
        'mpn_risk_map': {k: v for k, v in mpn_risk_map.items()}
    }

    # --- Persist Results to Database ---
    try:
        report = ReportAnalysis.objects.filter(user=request.user, report_file_path=file_name).first()
        if report:
            report.analysis_results = {
                'newData': new_data,
                'cpnRiskMap': cpn_risk_map,
                'dashboard': dashboard,
                'chip_counts': {
                    'high': sum(1 for r in new_data if r[-1] == 'High Risk'),
                    'low': sum(1 for r in new_data if r[-1] == 'Low Risk'),
                    'none': sum(1 for r in new_data if r[-1] == 'No Risk')
                },
                'columns': original_columns
            }
            report.high_risk_count = dashboard['cpn_high']
            report.low_risk_count = dashboard['cpn_low']
            report.no_risk_count = dashboard['cpn_no']
            report.is_analyzed = True
            report.save()
    except Exception as e:
        print(f"Error saving analysis to DB: {e}")
    
    return JsonResponse({
        'newData': new_data, 'columns': original_columns,
        'highRiskResults': hr_res, 'noRiskResults': nr_res, 'lowRiskResults': lr_res,
        'unmatchedUploads': unmatched_uploads, 'cpnRiskMap': cpn_risk_map,
        'dashboard': dashboard,
        'chip_counts': {
            'high': sum(1 for r in new_data if r[-1] == 'High Risk'),
            'low': sum(1 for r in new_data if r[-1] == 'Low Risk'),
            'none': sum(1 for r in new_data if r[-1] == 'No Risk')
        },
    })
