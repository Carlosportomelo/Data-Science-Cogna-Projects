import pandas as pd

P='outputs/visão resumida/hubspot_resumido_Dash.xlsx'
try:
    df=pd.read_excel(P, sheet_name='Visao_Granular_Final')
except Exception:
    df=pd.read_excel(P)

# parse dates
df['Data_Fechamento'] = pd.to_datetime(df['Data_Fechamento'], errors='coerce')
# ensure Matriculas numeric
if 'Matriculas' in df.columns:
    df['Matriculas'] = pd.to_numeric(df['Matriculas'], errors='coerce').fillna(0)
else:
    df['Matriculas'] = 0

# masks
mask_jan26 = (df['Data_Fechamento'].dt.year == 2026) & (df['Data_Fechamento'].dt.month == 1)
df_jan = df.loc[mask_jan26].copy()

# methods
sum_matriculas = int(df_jan['Matriculas'].sum())
rows_eq1 = int((df_jan['Matriculas'] == 1).sum())
unique_leads_matcol = int(df_jan.loc[df_jan['Matriculas']>0, 'Lead_Key'].nunique())

# status-based
status_mask = df_jan['Status_Principal'].astype(str).str.upper().str.contains('MATR', na=False)
unique_by_status = int(df_jan.loc[status_mask & df_jan['Data_Fechamento'].notna(), 'Lead_Key'].nunique())

print('RESULTADOS Jan/2026:')
print(f'  sum_matriculas (soma coluna): {sum_matriculas}')
print(f'  rows_eq1 (linhas Matriculas==1): {rows_eq1}')
print(f'  unique_leads_matcol (Lead_Key únicos com Matriculas>0): {unique_leads_matcol}')
print(f'  unique_by_status (Status contém "MATR" + Data_Fechamento): {unique_by_status}')

# breakdown by Origem_Principal for the status-based set
by_channel = df_jan.loc[status_mask & df_jan['Data_Fechamento'].notna()].groupby(df_jan['Origem_Principal'].fillna('Sem Canal'))['Lead_Key'].nunique()
print('\nPor canal (unique leads pela regra status+data):')
for k,v in by_channel.items():
    print(f'  {k}: {int(v)}')
