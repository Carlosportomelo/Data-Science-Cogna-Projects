import pandas as pd

print("Lendo datas com Sem Lead...")
# Usar skiprows e read_excel com engine otimizado
df = pd.read_excel(r'outputs\visão resumida\hubspot_resumido_Dash.xlsx', 
                    sheet_name='Visao_Granular_Final',
                    usecols=['Data', 'Status_Principal'],
                    engine='openpyxl')

sem_lead = df[df['Status_Principal'] == 'Sem Lead'].copy()
sem_lead['Data'] = pd.to_datetime(sem_lead['Data']).dt.normalize()
sem_lead = sem_lead[(sem_lead['Data'] >= '2025-10-01') & 
                    (sem_lead['Data'] <= '2026-03-02')]

print(f"Total registros 'Sem Lead': {len(sem_lead)}")

# Pegar as datas únicas
datas_com_sem_lead = sem_lead['Data'].unique()
print(f"Total DATAS com 'Sem Lead': {len(datas_com_sem_lead)}\n")

# Carregar investimentos
inv = pd.read_csv('output/days_invest_no_leads.csv')
inv['date'] = pd.to_datetime(inv['date']).dt.normalize()
inv = inv.rename(columns={'date': 'Data', 'meta_investimento': 'Meta', 'google_investimento': 'Google'})

# Filtrar investimentos apenas nos dias com "Sem Lead"
inv_com_sem_lead = inv[inv['Data'].isin(datas_com_sem_lead)].copy()

print(f"{'Data':<14} {'Meta':<15} {'Google':<15} {'Total':<15}")
print('=' * 65)

total_meta = 0
total_google = 0

for _, r in inv_com_sem_lead.sort_values('Data').iterrows():
    total = r['Meta'] + r['Google']
    total_meta += r['Meta']
    total_google += r['Google']
    destaque = ' ⚠️' if total > 5000 else ''
    
    print(f"{r['Data'].strftime('%d/%m/%Y'):<14} "
          f"R$ {r['Meta']:>10,.2f}  "
          f"R$ {r['Google']:>10,.2f}  "
          f"R$ {total:>10,.2f}{destaque}")

print('=' * 65)
print(f"Meta (Social Pago):     R$ {total_meta:>12,.2f}")
print(f"Google (Pesquisa Paga): R$ {total_google:>12,.2f}")
print(f"TOTAL em dias com Sem Lead: R$ {total_meta + total_google:>12,.2f}")
