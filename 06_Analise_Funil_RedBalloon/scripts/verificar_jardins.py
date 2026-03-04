import pandas as pd
from datetime import datetime, timedelta

df = pd.read_csv(r'C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_ATUAL.csv')
df['Data de criação'] = pd.to_datetime(df['Data de criação'])

# Ciclo 26.1: Out/2025 a Mar/2026
df_ciclo = df[
    (df['Data de criação'] >= '2025-10-01') & 
    (df['Data de criação'] <= '2026-03-31')
].copy()

# D-1 = idade ≤1 dia na data de extração (12/Fev/2026 - HOJE)
data_extracao = datetime(2026, 2, 12)
df_ciclo['idade_dias'] = (data_extracao - df_ciclo['Data de criação']).dt.days

df_d1_itaim = df_ciclo[
    (df_ciclo['idade_dias'] <= 1) &
    (df_ciclo['Unidade Desejada'] == 'ITAIM BIBI')
]

print(f'Leads Itaim Bibi D-1 (no ciclo 26.1): {len(df_d1_itaim)}')
print(f'Com atividade (>0): {(df_d1_itaim["Número de atividades de vendas"] > 0).sum()}')
print(f'SEM atividade (=0): {(df_d1_itaim["Número de atividades de vendas"] == 0).sum()}')
print('\nDistribuição de atividades:')
print(df_d1_itaim['Número de atividades de vendas'].value_counts().sort_index())

print('\n' + '='*100)
print('LEADS SEM ATIVIDADE (0):')
print('='*100)
sem_ativ = df_d1_itaim[df_d1_itaim['Número de atividades de vendas'] == 0]
if len(sem_ativ) > 0:
    colunas = ['Nome do negócio', 'Data de criação', 'Fonte original do tráfego', 
               'Etapa do negócio', 'Número de atividades de vendas', 'idade_dias']
    print(sem_ativ[colunas].to_string(index=False))
else:
    print('(Nenhum lead sem atividade encontrado)')


