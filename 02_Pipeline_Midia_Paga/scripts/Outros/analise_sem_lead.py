import pandas as pd
from pathlib import Path

# Arquivo resumido
file_path = Path('c:/Users/a483650/Projetos/analise_performance_midiapaga/outputs/visão resumida/hubspot_resumido_Dash.xlsx')
df = pd.read_excel(file_path, sheet_name='Blend_Agregado_Dash')

print('='*70)
print('COLUNAS DO DATAFRAME')
print('='*70)
print(df.columns.tolist())
print(df.head())

print('='*70)
print('ANALISE: INVESTIMENTO SEM LEAD')
print('='*70)

# Filtrar registros com 'Sem Lead' - verificar qual coluna identifica isso
if 'Status_Principal' in df.columns:
    df_sem_lead = df[df['Status_Principal'] == 'Sem Lead']
elif 'Unidade_Desejada' in df.columns:
    df_sem_lead = df[df['Unidade_Desejada'] == 'Sem Lead']
else:
    print("Colunas disponíveis:", df.columns.tolist())
    df_sem_lead = pd.DataFrame()

print(f'\nTotal de registros SEM LEAD: {len(df_sem_lead)}')
print(f'Investimento total SEM LEAD: R$ {df_sem_lead["Investimento"].sum():,.2f}')

# Breakdown por canal
print(f'\nPor Canal:')
if len(df_sem_lead) > 0:
    for canal in sorted(df_sem_lead['Canal'].unique()):
        inv = df_sem_lead[df_sem_lead['Canal'] == canal]['Investimento'].sum()
        count = len(df_sem_lead[df_sem_lead['Canal'] == canal])
        print(f'  - {canal}: R$ {inv:,.2f} ({count} dias)')

# Investimento total geral para comparação
print(f'\n' + '='*70)
print('COMPARACAO: INVESTIMENTO TOTAL vs SEM LEAD')
print('='*70)

inv_total = df['Investimento'].sum()
inv_com_lead = df[df['Unidade_Desejada'] != 'Sem Lead']['Investimento'].sum()
inv_sem_lead_total = df_sem_lead['Investimento'].sum()
pct_sem_lead = (inv_sem_lead_total / inv_total * 100) if inv_total > 0 else 0

print(f'\nInvestimento Total:     R$ {inv_total:,.2f}')
print(f'Com Lead Capturado:     R$ {inv_com_lead:,.2f}')
print(f'SEM Lead Capturado:     R$ {inv_sem_lead_total:,.2f}')
print(f'Percentual sem Lead: {pct_sem_lead:.2f}%')
