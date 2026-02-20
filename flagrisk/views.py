import pandas as pd
from django.shortcuts import render
from django.core.files.storage import FileSystemStorage
from django.conf import settings
import os

def index(request):
    try:
        if request.method == 'POST' and request.FILES.get('excel_file', None):
            file_obj = request.FILES['excel_file']
            fs = FileSystemStorage()

            fixed_filename = "uploaded_file.xlsx"
            file_path = os.path.join(settings.MEDIA_ROOT, fixed_filename)

            if fs.exists(fixed_filename):
                fs.delete(fixed_filename)

            filename = fs.save(fixed_filename, file_obj)
            file_path = fs.path(filename)

            
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
                    
                    df = df.dropna(how='all')
                    
                    df = df.loc[:, df.columns.notna()]
                    
                    cols_to_keep = [col for col in df.columns if not str(col).startswith('Unnamed')]
                    df = df[cols_to_keep]
                    
                    df = df.reset_index(drop=True)
                
                df = df.fillna('')
                
                columns = df.columns.tolist()
                data = df.values.tolist()
                
         
                
                context = {
                    'data': data,
                    'columns': columns,
                    'is_uploaded': True,
                    'filename': filename,
                }
                return render(request, 'flagrisk/index.html', context)
            except Exception as e:
                return render(request, 'flagrisk/index.html', {'error': f"Error reading file: {str(e)}"})
            
        return render(request, 'flagrisk/index.html')
    except Exception as e:
         return render(request, 'flagrisk/index.html', {'error': f"Error parsing request: {str(e)}"})
