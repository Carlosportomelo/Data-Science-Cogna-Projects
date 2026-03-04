import pandas as pd
from datetime import datetime

df = pd.read_csv(r'C:\Users\a483650\Projetos\06_Analise_Funil_RedBalloon\data\hubspot_leads_redballoon.csv')

print("=" * 80)
print("ANÁLISE DO CAMPO 'Data da última atividade'")
print("=" * 80)

print(f"\nTotal de registros: {len(df):,}")
print(f"Registros com 'Data da última atividade' preenchida: {df['Data da última atividade'].notna().sum():,}")
print(f"Registros sem data: {df['Data da última atividade'].isna().sum():,}")

# Ver exemplos de datas preenchidas
print("\n" + "-" * 80)
print("Exemplos de registros COM data da última atividade:")
print("-" * 80)
df_com_data = df[df['Data da última atividade'].notna()][['Record ID', 'Nome do negócio', 'Data de criação', 'Número de atividades de vendas', 'Data da última atividade']].head(10)
print(df_com_data.to_string())

# Converter e ver distribuição
df['Data da última atividade'] = pd.to_datetime(df['Data da última atividade'], errors='coerce')
df['Data de criação'] = pd.to_datetime(df['Data de criação'])

# Ver quantos leads têm atividade recente
hoje = datetime.now()
df_com_atividade = df[df['Data da última atividade'].notna()].copy()

if len(df_com_atividade) > 0:
    df_com_atividade['dias_desde_ultima_atividade'] = (hoje - df_com_atividade['Data da última atividade']).dt.days
    
    print("\n" + "=" * 80)
    print("DISTRIBUIÇÃO DE RECÊNCIA DA ÚLTIMA ATIVIDADE:")
    print("=" * 80)
    print(f"Atividade D-1 (≤1 dia):     {(df_com_atividade['dias_desde_ultima_atividade'] <= 1).sum():,} leads")
    print(f"Atividade < 7 dias:         {(df_com_atividade['dias_desde_ultima_atividade'] < 7).sum():,} leads")
    print(f"Atividade < 30 dias:        {(df_com_atividade['dias_desde_ultima_atividade'] < 30).sum():,} leads")
    print(f"Atividade < 60 dias:        {(df_com_atividade['dias_desde_ultima_atividade'] < 60).sum():,} leads")
    print(f"Atividade < 90 dias:        {(df_com_atividade['dias_desde_ultima_atividade'] < 90).sum():,} leads")
    print(f"Atividade > 90 dias:        {(df_com_atividade['dias_desde_ultima_atividade'] >= 90).sum():,} leads")
