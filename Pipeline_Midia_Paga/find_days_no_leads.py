import pandas as pd
from pathlib import Path

# Investimentos
meta = pd.read_excel('outputs/meta_dataset_dashboard.xlsx', sheet_name='Meta_YoY', usecols=['Data', 'Investimento'])
google = pd.read_excel('outputs/google_dashboard.xlsx', sheet_name='Google_YoY', usecols=['Data', 'Investimento_Google'])

meta['Data'] = pd.to_datetime(meta['Data']).dt.normalize()
google['Data'] = pd.to_datetime(google['Data']).dt.normalize()

inv = pd.merge(meta, google, on='Data', how='outer').fillna(0)
inv['Total'] = inv['Investimento'] + inv['Investimento_Google']
inv = inv[(inv['Data'] >= '2025-10-01') & (inv['Data'] <= '2026-03-02') & (inv['Total'] > 0)]

# Leads
hub = pd.read_excel('outputs/visão resumida/hubspot_resumido_Dash.xlsx', sheet_name='Visao_Granular_Final', usecols=['Data', 'Origem_Principal'])
hub['Data'] = pd.to_datetime(hub['Data']).dt.normalize()
hub = hub[(hub['Data'] >= '2025-10-01') & (hub['Data'] <= '2026-03-02') & (hub['Origem_Principal'].isin(['Social Pago', 'Pesquisa Paga']))]
leads = hub.groupby('Data').size().reset_index(name='Leads')

# Cruzamento
result = pd.merge(inv[['Data', 'Total']], leads, on='Data', how='left')
sem_leads = result[result['Leads'].isna()].sort_values('Data')

print(f'Total: {len(sem_leads)} dias com investimento mas sem leads\n')
for _, r in sem_leads.iterrows():
    print(f"{r['Data'].strftime('%d/%m/%Y')} - R$ {r['Total']:,.2f}")
