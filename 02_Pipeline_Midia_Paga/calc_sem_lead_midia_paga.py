import pandas as pd

print("Lendo dados...")
df = pd.read_excel(r'outputs\visão resumida\hubspot_resumido_Dash.xlsx', 
                    sheet_name='Visao_Granular_Final',
                    usecols=['Data', 'Status_Principal', 'Midia_Paga', 'Origem_Principal'],
                    engine='openpyxl')

# Conversão de data
df['Data'] = pd.to_datetime(df['Data']).dt.normalize()

# Filtrar por Status_Principal = "Sem Lead" e período OUT/2025 A MAR/2026
sem_lead = df[(df['Status_Principal'] == 'Sem Lead') &
              (df['Data'] >= '2025-10-01') & 
              (df['Data'] <= '2026-03-31')].copy()

print(f"Total registros 'Sem Lead' (out 2025 - mar 2026): {len(sem_lead)}\n")

# Ordenar por data e origem
sem_lead_sort = sem_lead.sort_values(['Data', 'Origem_Principal'])

print(f"{'Data':<12} {'Canal':<20} {'Midia_Paga':<15}")
print('=' * 50)

total_geral = 0
for _, row in sem_lead_sort.iterrows():
    data_str = row['Data'].strftime('%d/%m/%Y')
    canal = row['Origem_Principal']
    valor = row['Midia_Paga']
    total_geral += valor
    
    print(f"{data_str:<12} {canal:<20} R$ {valor:>12,.2f}")

print('=' * 50)
print(f"\nTOTAL DE MIDIA_PAGA (Sem Lead): R$ {total_geral:,.2f}")
