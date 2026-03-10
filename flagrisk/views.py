import pandas as pd
from django.shortcuts import render
from django.core.files.storage import FileSystemStorage
from django.conf import settings
import os
import datetime
import glob

def ensure_limit(reports_dir):
    all_hist = glob.glob(os.path.join(reports_dir, "report_*.xlsx"))
    all_hist.sort(key=os.path.getmtime, reverse=True)
    if len(all_hist) > 5:
        for old in all_hist[5:]:
            if os.path.isfile(old):
                os.remove(old)

def index(request):
    fs = FileSystemStorage()
    reports_dir = os.path.join(settings.MEDIA_ROOT, 'reports')
    if not os.path.exists(reports_dir):
        os.makedirs(reports_dir)

    # 1. Handle Clear History request
    if request.GET.get('clear') == '1':
        files = glob.glob(os.path.join(reports_dir, "*"))
        for f in files:
            if os.path.isfile(f):
                os.remove(f)
        main_file = os.path.join(settings.MEDIA_ROOT, "uploaded_file.xlsx")
        if os.path.exists(main_file):
            os.remove(main_file)
        from django.shortcuts import redirect
        return redirect('/')

    # 2. Strict Rotation
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
        if (request.method == 'POST' and request.FILES.get('excel_file')) or file_path:
            if not file_path:
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

                ensure_limit(reports_dir)

            try:  
                df_raw = pd.read_excel(file_path, header=None)
                header_row_idx = None
                for idx, row in df_raw.iterrows():
                    row_values = [str(cell).lower().strip() if pd.notna(cell) else '' for cell in row]
                    key_columns = ['part number', 'manufacturer', 'cpn', 'georisk']
                    matches = sum(1 for col in key_columns if any(col in val for val in row_values))
                    if matches >= 3:
                        header_row_idx = idx
                        break
                
                if header_row_idx is None:
                    df = pd.read_excel(file_path)
                else:
                    df = pd.read_excel(file_path, header=header_row_idx)
                
                df = df.dropna(how='all').loc[:, df.columns.notna()]
                df = df[[c for c in df.columns if not str(c).startswith('Unnamed')]].reset_index(drop=True).fillna('')
                columns = df.columns.tolist()
                data = df.values.tolist()
                
                recent_list = []
                all_hist = glob.glob(os.path.join(reports_dir, "report_*.xlsx"))
                all_hist.sort(key=os.path.getmtime, reverse=True)
                for f in all_hist[:5]:
                    mtime = os.path.getmtime(f)
                    dt = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
                    fname = os.path.basename(f)
                    display_name = fname[22:] if fname.startswith("report_") and len(fname) > 22 else fname
                    recent_list.append({'name': display_name, 'file': fname, 'date': dt})

                return render(request, 'flagrisk/index.html', {
                    'data': data,
                    'columns': columns,
                    'is_uploaded': True,
                    'filename': original_name,
                    'recent_reports': recent_list
                })
            except Exception as e:
                return render(request, 'flagrisk/index.html', {'error': f"Error parsing Excel: {str(e)}"})
            
        recent_list = []
        all_hist = glob.glob(os.path.join(reports_dir, "report_*.xlsx"))
        all_hist.sort(key=os.path.getmtime, reverse=True)
        for f in all_hist[:5]:
            mtime = os.path.getmtime(f)
            dt = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
            fname = os.path.basename(f)
            display_name = fname[22:] if fname.startswith("report_") and len(fname) > 22 else fname
            recent_list.append({'name': display_name, 'file': fname, 'date': dt})

        return render(request, 'flagrisk/index.html', {'recent_reports': recent_list})
        
    except Exception as e:
         return render(request, 'flagrisk/index.html', {'error': f"General error: {str(e)}"})
