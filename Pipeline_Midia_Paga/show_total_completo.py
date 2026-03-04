import pandas as pd

print("Carregando dados...")

# 1. Ler dias com investimento mas zero leads
inv_csv = pd.read_csv('output/days_invest_no_leads.csv')
inv_csv['date'] = pd.to_datetime(inv_csv['date']).dt.normalize()
datas_investimento_zero_leads = set(inv_csv['date'].unique())
print(f"Dias com investimento mas zero leads: {len(datas_investimento_zero_leads)}")

# 2. Ler dias com status "Sem Lead"
df = pd.read_excel(r'outputs\visão resumida\hubspot_resumido_Dash.xlsx', 
                    sheet_name='Visao_Granular_Final',
                    usecols=['Data', 'Status_Principal'],
                    engine='openpyxl')

sem_lead = df[df['Status_Principal'] == 'Sem Lead'].copy()
sem_lead['Data'] = pd.to_datetime(sem_lead['Data']).dt.normalize()
sem_lead = sem_lead[(sem_lead['Data'] >= '2025-10-01') & (sem_lead['Data'] <= '2026-03-02')]
datas_sem_lead = set(sem_lead['Data'].unique())
print(f"Dias com status 'Sem Lead': {len(datas_sem_lead)}")

# 3. Combinar todas as datas problemáticas
todas_datas_problema = datas_investimento_zero_leads.union(datas_sem_lead)
print(f"Total de datas únicas com problema: {len(todas_datas_problema)}\n")

# 4. Ler investimentos completos e filtrar por essas datas
inv = pd.read_csv('output/days_invest_no_leads.csv')
inv['date'] = pd.to_datetime(inv['date']).dt.normalize()
inv = inv.rename(columns={'date': 'Data', 'meta_investimento': 'Meta', 'google_investimento': 'Google'})
inv = inv[['Data', 'Meta', 'Google']]

# Filtrar apenas as datas problemáticas
inv_total = inv[inv['Data'].isin(todas_datas_problema)].copy()

print(f"{'Data':<14} {'Meta':<15} {'Google':<15} {'Total':<15} Tipo")
print('=' * 85)

total_meta = 0
total_google = 0

for _, r in inv_total.sort_values('Data').iterrows():
    total = r['Meta'] + r['Google']
    total_meta += r['Meta']
    total_google += r['Google']
    
    # Determinar tipo
    eh_zero = r['Data'] in datas_investimento_zero_leads
    eh_sem_lead = r['Data'] in datas_sem_lead
    
    if eh_zero and eh_sem_lead:
        tipo = "Zero Leads + Sem Lead"
    elif eh_zero:
        tipo = "Zero Leads"
    else:
        tipo = "Sem Lead"
    
    destaque = ' ⚠️' if total > 5000 else ''
    
    print(f"{r['Data'].strftime('%d/%m/%Y'):<14} "
          f"R$ {r['Meta']:>10,.2f}  "
          f"R$ {r['Google']:>10,.2f}  "
          f"R$ {total:>10,.2f}{destaque}  {tipo}")

print('=' * 85)
print(f"\nMeta (Social Pago):     R$ {total_meta:>12,.2f}")
print(f"Google (Pesquisa Paga): R$ {total_google:>12,.2f}")
print(f"TOTAL DESPERDIÇADO:     R$ {total_meta + total_google:>12,.2f}")
