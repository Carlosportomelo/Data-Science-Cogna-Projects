import pandas as pd
from pathlib import Path

# Arquivo resumido
file_path = Path('c:/Users/a483650/Projetos/analise_performance_midiapaga/outputs/visão resumida/hubspot_resumido_Dash.xlsx')
df = pd.read_excel(file_path, sheet_name='Blend_Agregado_Dash')

# Converter data para datetime
df['Data_Criacao'] = pd.to_datetime(df['Data_Criacao'])

# Datas de Social Pago sem lead
datas_social = [
    '2026-01-28',
    '2026-01-29',
    '2026-01-30',
    '2026-01-31',
    '2026-02-01'
]

# Datas de Pesquisa Paga sem lead
datas_pesquisa = [
    '2026-01-10',
    '2026-01-25',
    '2026-01-27',
    '2026-01-28',
    '2026-01-29',
    '2026-01-30',
    '2026-01-31',
    '2026-02-01',
    '2026-02-02'
]

print('='*70)
print('INVESTIMENTO POR DATAS SEM LEAD (JANEIRO/FEVEREIRO 2026)')
print('='*70)

# Filtrar dados de Social Pago
df_social = df[(df['Canal'] == 'Social Pago') & (df['Data_Criacao'].dt.date.astype(str).isin(datas_social))]
inv_social = df_social['Investimento'].sum()

print(f'\nSOCIAL PAGO (5 dias):')
print(f'  Datas: {", ".join(datas_social)}')
print(f'  Total: R$ {inv_social:,.2f}')
print(f'  Registros: {len(df_social)}')

# Filtrar dados de Pesquisa Paga
df_pesquisa = df[(df['Canal'] == 'Pesquisa Paga') & (df['Data_Criacao'].dt.date.astype(str).isin(datas_pesquisa))]
inv_pesquisa = df_pesquisa['Investimento'].sum()

print(f'\nPESQUISA PAGA (9 dias):')
print(f'  Datas: {", ".join(datas_pesquisa)}')
print(f'  Total: R$ {inv_pesquisa:,.2f}')
print(f'  Registros: {len(df_pesquisa)}')

print(f'\n' + '='*70)
print(f'INVESTIMENTO TOTAL SEM LEAD (ESSAS DATAS): R$ {inv_social + inv_pesquisa:,.2f}')
print('='*70)

# Mostrar os investimentos por data para Social
print(f'\nDETALHE - SOCIAL PAGO:')
for data in datas_social:
    inv = df_social[df_social['Data_Criacao'].dt.date.astype(str) == data]['Investimento'].sum()
    if inv > 0:
        print(f'  {data}: R$ {inv:,.2f}')

# Mostrar os investimentos por data para Pesquisa
print(f'\nDETALHE - PESQUISA PAGA:')
for data in datas_pesquisa:
    inv = df_pesquisa[df_pesquisa['Data_Criacao'].dt.date.astype(str) == data]['Investimento'].sum()
    if inv > 0:
        print(f'  {data}: R$ {inv:,.2f}')
