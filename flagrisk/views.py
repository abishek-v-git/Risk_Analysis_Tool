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
            filename = fs.save(file_obj.name, file_obj)
            file_path = fs.path(filename)
            
            try:
                # Read the excel file without headers first to find the Parts Grid
                df_raw = pd.read_excel(file_path, header=None)
                
                # Find the row containing the actual Parts Grid column headers
                # We look for a row that has multiple key column names
                header_row_idx = None
                for idx, row in df_raw.iterrows():
                    # Convert row to list of strings (lowercase)
                    row_values = [str(cell).lower().strip() if pd.notna(cell) else '' for cell in row]
                    
                    # Check if this row contains multiple key column headers
                    key_columns = ['part number', 'manufacturer', 'cpn', 'georisk']
                    matches = sum(1 for col in key_columns if any(col in val for val in row_values))
                    
                    # If we find at least 3 of the key columns, this is likely our header row
                    if matches >= 3:
                        header_row_idx = idx
                        break
                
                if header_row_idx is None:
                    # If we can't find the header, try default behavior
                    df = pd.read_excel(file_path)
                else:
                    # Read again with the correct header row
                    df = pd.read_excel(file_path, header=header_row_idx)
                    
                    # Remove any completely empty rows
                    df = df.dropna(how='all')
                    
                    # Remove columns that are completely empty or have no meaningful name
                    df = df.loc[:, df.columns.notna()]
                    
                    # Remove unnamed columns
                    cols_to_keep = [col for col in df.columns if not str(col).startswith('Unnamed')]
                    df = df[cols_to_keep]
                    
                    # Reset index
                    df = df.reset_index(drop=True)
                
                # Replace NaN with empty string for display
                df = df.fillna('')
                
                # Get columns and data
                columns = df.columns.tolist()
                data = df.values.tolist()
                
                # Optionally delete the file after processing
                # os.remove(file_path) 
                
                context = {
                    'data': data,
                    'columns': columns,
                    'is_uploaded': True,
                    'filename': filename,
                }
                return render(request, 'flagrisk/index.html', context)
            except Exception as e:
                # If error reading file
                return render(request, 'flagrisk/index.html', {'error': f"Error reading file: {str(e)}"})
            
        return render(request, 'flagrisk/index.html')
    except Exception as e:
         return render(request, 'flagrisk/index.html', {'error': f"Error parsing request: {str(e)}"})
