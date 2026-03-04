import pandas as pd
from datetime import datetime, timedelta

df = pd.read_csv(r'C:\Users\a483650\Projetos\06_Analise_Funil_RedBalloon\data\hubspot_leads_redballoon.csv')
df['Data de criação'] = pd.to_datetime(df['Data de criação'])

# Filtrar ciclo 26.1
df_ciclo_26_1 = df[(df['Data de criação'] >= '2025-10-01') & (df['Data de criação'] <= '2026-03-31')].copy()

# Calcular idade
hoje = datetime.now()
df_ciclo_26_1['idade_dias'] = (hoje - df_ciclo_26_1['Data de criação']).dt.days

# Classificar
def classificar_idade(dias):
    if pd.isna(dias):
        return 'Sem data'
    elif dias <= 1:
        return 'D-1'
    elif dias < 7:
        return '< 7 dias'
    elif dias < 30:
        return '< 30 dias'
    elif dias < 60:
        return '< 60 dias'
    elif dias < 90:
        return '< 90 dias'
    else:
        return '> 90 dias'

df_ciclo_26_1['faixa_idade'] = df_ciclo_26_1['idade_dias'].apply(classificar_idade)
df_ciclo_26_1['canal'] = df_ciclo_26_1['Fonte original do tráfego'].fillna('Não informado')

# Distribuição por canal e faixa
distribuicao = df_ciclo_26_1.groupby(['canal', 'faixa_idade']).size().unstack(fill_value=0)
print("DISTRIBUIÇÃO POR CANAL E FAIXA:")
print(distribuicao)

# Total por canal
total_canal = df_ciclo_26_1.groupby('canal').size()
print("\n\nTOTAL POR CANAL:")
print(total_canal)

# Testar uma faixa específica
print("\n\nLEADS D-1:")
leads_d1 = df_ciclo_26_1[df_ciclo_26_1['faixa_idade'] == 'D-1']
print(f"Total D-1: {len(leads_d1)}")
print("\nPor canal:")
print(leads_d1.groupby('canal').size())

# Atendidos D-1
print("\n\nATENDIDOS D-1:")
atendidos_d1 = leads_d1[leads_d1['Número de atividades de vendas'] > 0]
print(f"Total atendidos: {len(atendidos_d1)}")
print("\nPor canal:")
print(atendidos_d1.groupby('canal').size())
