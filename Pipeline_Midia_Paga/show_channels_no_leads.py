import pandas as pd

print("Carregando investimentos por canal...")
# Carregar o CSV que já tem os investimentos por canal
inv = pd.read_csv('output/days_invest_no_leads.csv')
inv['date'] = pd.to_datetime(inv['date'])
inv = inv.rename(columns={'date': 'Data', 'meta_investimento': 'Meta', 'google_investimento': 'Google'})
inv = inv[['Data', 'Meta', 'Google']]

print("Carregando leads por canal (Meta)...")
# Ler apenas dados Meta do dashboard específico
meta_df = pd.read_excel('outputs/meta_dataset_dashboard.xlsx', 
                        sheet_name='Meta_YoY',
                        usecols=['Data'])
meta_df['Data'] = pd.to_datetime(meta_df['Data']).dt.normalize()
meta_df = meta_df[(meta_df['Data'] >= '2025-10-01') & (meta_df['Data'] <= '2026-03-02')]

print("Carregando leads por canal (Google)...")
# Ler apenas dados Google do dashboard específico
google_df = pd.read_excel('outputs/google_dashboard.xlsx',
                          sheet_name='Google_YoY', 
                          usecols=['Data'])
google_df['Data'] = pd.to_datetime(google_df['Data']).dt.normalize()
google_df = google_df[(google_df['Data'] >= '2025-10-01') & (google_df['Data'] <= '2026-03-02')]

print("Processando dados...")

# Filtrar período
blend = blend[(blend['Data'] >= '2025-10-01') & (blend['Data'] <= '2026-03-02')]

# Contar leads por canal por dia
leads_meta = blend[blend['Origem_Principal'] == 'Social Pago'].groupby('Data').size().reset_index(name='Leads_Meta')
leads_google = blend[blend['Origem_Principal'] == 'Pesquisa Paga'].groupby('Data').size().reset_index(name='Leads_Google')

# Merge com investimentos
result = inv.copy()
result = pd.merge(result, leads_meta, on='Data', how='left')
result = pd.merge(result, leads_google, on='Data', how='left')
result['Leads_Meta'] = result['Leads_Meta'].fillna(0).astype(int)
result['Leads_Google'] = result['Leads_Google'].fillna(0).astype(int)

# Identificar canais sem leads mas com investimento
def canais_sem_leads(row):
    canais = []
    if row['Meta'] > 0 and row['Leads_Meta'] == 0:
        canais.append('Meta (Social Pago)')
    if row['Google'] > 0 and row['Leads_Google'] == 0:
        canais.append('Google (Pesquisa Paga)')
    return ' + '.join(canais) if canais else ''

result['Canais_Sem_Leads'] = result.apply(canais_sem_leads, axis=1)

# Filtrar apenas dias com pelo menos um canal sem leads
sem_leads = result[result['Canais_Sem_Leads'] != ''].sort_values('Data')

print(f'\nTotal: {len(sem_leads)} dias com investimento sem leads (out/2025 a 02/03/2026)\n')
print(f"{'Data':<14} {'Meta':<15} {'Google':<15} {'Total':<15} Canais Sem Leads")
print('=' * 90)

total_meta_desperdiçado = 0
total_google_desperdiçado = 0

for _, r in sem_leads.iterrows():
    total = r['Meta'] + r['Google']
    meta_desperdiçado = r['Meta'] if r['Meta'] > 0 and r['Leads_Meta'] == 0 else 0
    google_desperdiçado = r['Google'] if r['Google'] > 0 and r['Leads_Google'] == 0 else 0
    
    total_meta_desperdiçado += meta_desperdiçado
    total_google_desperdiçado += google_desperdiçado
    
    destaque = ' ⚠️' if total > 5000 else ''
    
    print(f"{r['Data'].strftime('%d/%m/%Y'):<14} "
          f"R$ {r['Meta']:>10,.2f}  "
          f"R$ {r['Google']:>10,.2f}  "
          f"R$ {total:>10,.2f}{destaque}  "
          f"{r['Canais_Sem_Leads']}")

print('=' * 90)
print(f"\nResumo do investimento desperdiçado:")
print(f"Meta (Social Pago):        R$ {total_meta_desperdiçado:>12,.2f}")
print(f"Google (Pesquisa Paga):    R$ {total_google_desperdiçado:>12,.2f}")
print(f"TOTAL DESPERDIÇADO:        R$ {total_meta_desperdiçado + total_google_desperdiçado:>12,.2f}")
