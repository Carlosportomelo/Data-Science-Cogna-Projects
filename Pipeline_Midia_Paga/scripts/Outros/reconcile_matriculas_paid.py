import pandas as pd

P='outputs/visão resumida/hubspot_resumido_Dash.xlsx'
df=pd.read_excel(P, sheet_name='Visao_Granular_Final')

df['Data_Fechamento']=pd.to_datetime(df['Data_Fechamento'], errors='coerce')
if 'Matriculas' in df.columns:
    df['Matriculas'] = pd.to_numeric(df['Matriculas'], errors='coerce').fillna(0)
else:
    df['Matriculas'] = 0

mask_jan = (df['Data_Fechamento'].dt.year == 2026) & (df['Data_Fechamento'].dt.month == 1)
df_jan = df.loc[mask_jan].copy()

paid = ['Social Pago', 'Pesquisa Paga']
# by two heuristics: Matriculas>0 and Status contains 'MATR'
paid_mask = df_jan['Origem_Principal'].astype(str).isin(paid)
mat_mask = df_jan['Matriculas'] > 0
status_mask = df_jan['Status_Principal'].astype(str).str.upper().str.contains('MATR', na=False)

paid_mat = df_jan[paid_mask & mat_mask]
paid_status = df_jan[paid_mask & status_mask]

print('Paid channels considered:', paid)
print('Count (Matriculas>0) for paid channels:', len(paid_mat))
print('Unique Lead_Key (Matriculas>0) for paid channels:', paid_mat['Lead_Key'].nunique())
print('\nLead_Key samples (Matriculas>0, paid):')
print(paid_mat['Lead_Key'].drop_duplicates().head(100).to_string(index=False))

print('\nCount (Status contains MATR) for paid channels:', len(paid_status))
print('Unique Lead_Key (Status MATR) for paid channels:', paid_status['Lead_Key'].nunique())
print('\nLead_Key samples (Status MATR, paid):')
print(paid_status['Lead_Key'].drop_duplicates().head(100).to_string(index=False))

# export lists for review
OUT='outputs/visão resumida/paid_matriculas_jan2026.csv'
paid_mat.to_csv(OUT, index=False, encoding='utf-8-sig')
print(f'Exported paid Matriculas rows to: {OUT}')
