import pandas as pd

print("Lendo arquivo...")
# Ler a aba Visao_Granular_Final
df = pd.read_excel(r'outputs\visão resumida\hubspot_resumido_Dash.xlsx', 
                    sheet_name='Visao_Granular_Final')

print(f"Total de registros: {len(df)}")
print(f"Colunas: {df.columns.tolist()}\n")

# Verificar o que tem em Status_Principal
print("Valores únicos em Status_Principal:")
print(df['Status_Principal'].value_counts())
print()

# Filtrar por "Sem Lead" (verificar exato)
sem_lead = df[df['Status_Principal'] == 'Sem Lead'].copy()
print(f"Total de registros 'Sem Lead': {len(sem_lead)}\n")

# Preparar para análise
sem_lead['Data'] = pd.to_datetime(sem_lead['Data']).dt.normalize()

# Filtrar período Oct 2025 - Mar 2 2026
sem_lead = sem_lead[(sem_lead['Data'] >= '2025-10-01') & (sem_lead['Data'] <= '2026-03-02')]

print(f"Total sem lead no período (out/2025 a 02/03/2026): {len(sem_lead)}\n")

# Carregar investimentos por canal
inv = pd.read_csv('output/days_invest_no_leads.csv')
inv['date'] = pd.to_datetime(inv['date']).dt.normalize()
inv = inv.rename(columns={'date': 'Data', 'meta_investimento': 'Meta', 'google_investimento': 'Google'})

# Agrupar sem_lead por data e canal
sem_lead_por_dia = sem_lead.groupby(['Data', 'Origem_Principal']).size().reset_index(name='sem_leads')

print("Sem leads por data e canal:")
print(sem_lead_por_dia.head(20))
print()

# Merge com investimentos
resultado = inv.merge(sem_lead_por_dia, on='Data', how='left')
resultado['sem_leads'] = resultado['sem_leads'].fillna(0).astype(int)

# Identificar o canal responsável pelo investimento sem lead
def get_canal_sem_lead(row):
    canais = []
    if row['Meta'] > 0:
        # Há investimento em Meta
        meta_rows = sem_lead_por_dia[(sem_lead_por_dia['Data'] == row['Data']) & 
                                     (sem_lead_por_dia['Origem_Principal'] == 'Social Pago')]
        if len(meta_rows) > 0:
            canais.append(f"Meta (Social Pago) - {meta_rows['sem_leads'].values[0]} sem leads")
    
    if row['Google'] > 0:
        # Há investimento em Google
        google_rows = sem_lead_por_dia[(sem_lead_por_dia['Data'] == row['Data']) & 
                                       (sem_lead_por_dia['Origem_Principal'] == 'Pesquisa Paga')]
        if len(google_rows) > 0:
            canais.append(f"Google (Pesquisa Paga) - {google_rows['sem_leads'].values[0]} sem leads")
    
    return ' | '.join(canais) if canais else ''

resultado['Canais_Sem_Lead'] = resultado.apply(get_canal_sem_lead, axis=1)

# Filtrar apenas dias com realmente sem leads
resultado_final = resultado[resultado['Canais_Sem_Lead'] != ''].sort_values('Data')

print(f"\nDias com investimento e registros 'sem lead':")
print(f"Total: {len(resultado_final)} dias\n")

print(f"{'Data':<14} {'Meta':<15} {'Google':<15} {'Total':<15} Canais", file=None)
print('=' * 100)

for _, r in resultado_final.iterrows():
    total = r['Meta'] + r['Google']
    destaque = ' ⚠️' if total > 5000 else ''
    
    print(f"{r['Data'].strftime('%d/%m/%Y'):<14} "
          f"R$ {r['Meta']:>10,.2f}  "
          f"R$ {r['Google']:>10,.2f}  "
          f"R$ {total:>10,.2f}{destaque}  "
          f"{r['Canais_Sem_Lead']}")

print('=' * 100)
print(f"Total investido em dias com 'sem lead': R$ {resultado_final['Meta'].sum() + resultado_final['Google'].sum():,.2f}")
