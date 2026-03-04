import pandas as pd
path = 'Outputs/Executive_Package_AllSheets_2025-12-11.xlsx'
try:
    sheets = pd.read_excel(path, sheet_name=None)
    for name in sheets.keys():
        print(name)
except Exception as e:
    print('ERROR', e)
