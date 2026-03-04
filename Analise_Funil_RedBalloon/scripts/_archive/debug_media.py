import pandas as pd
from datetime import datetime

df = pd.read_csv(r'C:\Users\a483650\Projetos\06_Analise_Funil_RedBalloon\data\hubspot_leads_redballoon.csv')
df['Data de criação'] = pd.to_datetime(df['Data de criação'])

# Filtrar ciclo 26.1
df_26_1 = df[(df['Data de criação'] >= '2025-10-01') & (df['Data de criação'] <= '2026-03-31')]

print(f"Total leads 26.1: {len(df_26_1)}")
print(f"\nEstatísticas da coluna 'Número de atividades de vendas':")
print(f"- Total de leads: {len(df_26_1)}")
print(f"- Com atividades (não nulo): {df_26_1['Número de atividades de vendas'].notna().sum()}")
print(f"- Sem atividades (nulo): {df_26_1['Número de atividades de vendas'].isna().sum()}")

print("\n" + "="*80)
print("ANÁLISE POR CANAL:")
print("="*80)

canais = ['Pesquisa paga', 'Social orgânico', 'Pesquisa orgânica', 'Tráfego direto']

for canal in canais:
    df_canal = df_26_1[df_26_1['Fonte original do tráfego'] == canal]
    print(f"\n{canal}:")
    print(f"  Total leads: {len(df_canal)}")
    print(f"  Com atividades: {df_canal['Número de atividades de vendas'].notna().sum()}")
    print(f"  Sem atividades (NaN): {df_canal['Número de atividades de vendas'].isna().sum()}")
    print(f"  Média (incluindo NaN): {df_canal['Número de atividades de vendas'].mean():.2f}")
    print(f"  Média (excluindo NaN): {df_canal['Número de atividades de vendas'].dropna().mean():.2f}")
    print(f"  Mediana: {df_canal['Número de atividades de vendas'].median():.2f}")
    print(f"  Distribuição:")
    print(df_canal['Número de atividades de vendas'].describe())
