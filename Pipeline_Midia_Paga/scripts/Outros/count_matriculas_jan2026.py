import pandas as pd

P='outputs/visão resumida/hubspot_resumido_Dash.xlsx'
try:
    df=pd.read_excel(P, sheet_name='Visao_Granular_Final')
except Exception:
    df=pd.read_excel(P)

if 'Data_Fechamento' not in df.columns:
    raise SystemExit('Coluna Data_Fechamento não encontrada')

# normalizar
df['Data_Fechamento']=pd.to_datetime(df['Data_Fechamento'], errors='coerce')
if 'Matriculas' not in df.columns:
    raise SystemExit('Coluna Matriculas não encontrada')

df['Matriculas']=pd.to_numeric(df['Matriculas'], errors='coerce').fillna(0)

mask=(df['Data_Fechamento'].dt.year==2026) & (df['Data_Fechamento'].dt.month==1)
dfj=df.loc[mask].copy()

# metrics
sum_matriculas = int(dfj['Matriculas'].sum())
rows_eq1 = int((dfj['Matriculas']==1).sum())
unique_leads = int(dfj.loc[dfj['Matriculas']>0, 'Lead_Key'].nunique())

by_channel = dfj.groupby(dfj['Origem_Principal'].fillna('Sem Canal'))['Matriculas'].sum()

print('SUM_MATRICULAS:', sum_matriculas)
print('ROWS_EQ_1:', rows_eq1)
print('UNIQUE_LEADS_WITH_MAT_GT0:', unique_leads)
print('\nBY_CHANNEL:')
for k,v in by_channel.items():
    print(f"  {k}: {int(v)}")

# sample
print('\nSAMPLE (matriculas>0):')
print(dfj.loc[dfj['Matriculas']>0, ['Lead_Key','Data_Fechamento','Matriculas','Origem_Principal']].head(20).to_string(index=False))
