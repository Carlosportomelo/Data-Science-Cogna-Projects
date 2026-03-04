import pandas as pd
from pathlib import Path

# Arquivo resumido
file_path = Path('c:/Users/a483650/Projetos/analise_performance_midiapaga/outputs/visão resumida/hubspot_resumido_Dash.xlsx')
df = pd.read_excel(file_path, sheet_name='Blend_Agregado_Dash')

# Converter data para datetime
df['Data_Criacao'] = pd.to_datetime(df['Data_Criacao'])
df['Ano'] = df['Data_Criacao'].dt.year
df['Mes'] = df['Data_Criacao'].dt.month

print('='*70)
print('INVESTIMENTO SEM LEAD POR ANO')
print('='*70)

# Filtrar registros com 'Sem Lead'
df_sem_lead = df[df['Unidade_Desejada'] == 'Sem Lead'].copy()

anos = sorted(df_sem_lead['Ano'].unique())
for ano in anos:
    df_ano = df_sem_lead[df_sem_lead['Ano'] == ano]
    inv_total = df_ano['Investimento'].sum()
    count = len(df_ano)
    print(f'\nAno {ano}: R$ {inv_total:,.2f} ({count} registros)')
    
    # Breakdown por canal
    for canal in sorted(df_ano['Canal'].unique()):
        inv_canal = df_ano[df_ano['Canal'] == canal]['Investimento'].sum()
        count_canal = len(df_ano[df_ano['Canal'] == canal])
        print(f'  - {canal}: R$ {inv_canal:,.2f} ({count_canal} dias)')

print('\n' + '='*70)
print('INVESTIMENTO SEM LEAD POR MÊS (2025)')
print('='*70)

df_2025 = df_sem_lead[df_sem_lead['Ano'] == 2025]
if len(df_2025) > 0:
    for mes in sorted(df_2025['Mes'].unique()):
        df_mes = df_2025[df_2025['Mes'] == mes]
        inv_mes = df_mes['Investimento'].sum()
        print(f'  {mes:02d}/2025: R$ {inv_mes:,.2f}')
else:
    print("Sem dados de 2025")

print('\n' + '='*70)
print('INVESTIMENTO SEM LEAD POR MÊS (2024)')
print('='*70)

df_2024 = df_sem_lead[df_sem_lead['Ano'] == 2024]
if len(df_2024) > 0:
    for mes in sorted(df_2024['Mes'].unique()):
        df_mes = df_2024[df_2024['Mes'] == mes]
        inv_mes = df_mes['Investimento'].sum()
        print(f'  {mes:02d}/2024: R$ {inv_mes:,.2f}')
else:
    print("Sem dados de 2024")
