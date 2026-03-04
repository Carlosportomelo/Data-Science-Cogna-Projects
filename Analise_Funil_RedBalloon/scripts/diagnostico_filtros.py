import pandas as pd

# Carregar dados
ARQUIVO_CSV = r'C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_ATUAL.csv'
df = pd.read_csv(ARQUIVO_CSV, sep=',')

# Converter colunas de data
df['Data de criação'] = pd.to_datetime(df['Data de criação'], errors='coerce')
df['Data da última atividade'] = pd.to_datetime(df['Data da última atividade'], errors='coerce')

# Definir ciclo 26.1
def definir_ciclo(data_criacao):
    if pd.isna(data_criacao):
        return None
    ano = data_criacao.year
    mes = data_criacao.month
    
    # Ciclo X.1 = Out/Nov/Dez do ano X-1 + Jan/Fev/Mar do ano X
    if mes >= 10:  # Out, Nov, Dez
        return f'{ano + 1}.1'
    elif mes <= 3:  # Jan, Fev, Mar
        return f'{ano}.1'
    elif mes >= 4 and mes <= 9:  # Abr-Set
        return f'{ano}.2'
    return None

df['ciclo'] = df['Data de criação'].apply(definir_ciclo)

print("\nVerificando distribuição de ciclos:")
print(df['ciclo'].value_counts().head(10))

df_ciclo_26_1 = df[df['ciclo'] == '2026.1'].copy()

print(f"Total de leads no ciclo 26.1: {len(df_ciclo_26_1):,}")

# Verificar correlação entre "Data da última atividade" e "Número de atividades"
print("\n" + "="*80)
print("ANÁLISE: Correlação entre Data da última atividade e Número de atividades")
print("="*80)

# Leads COM data de atividade válida
df_com_data = df_ciclo_26_1[
    (df_ciclo_26_1['Data da última atividade'].notna()) &
    (df_ciclo_26_1['Data da última atividade'] >= df_ciclo_26_1['Data de criação'])
].copy()

# Leads SEM data de atividade válida
df_sem_data = df_ciclo_26_1[
    (df_ciclo_26_1['Data da última atividade'].isna()) |
    (df_ciclo_26_1['Data da última atividade'] < df_ciclo_26_1['Data de criação'])
].copy()

print(f"\nLeads COM data de atividade válida: {len(df_com_data):,}")
print(f"Leads SEM data de atividade válida: {len(df_sem_data):,}")

# Nos leads COM data válida, quantos foram atendidos?
atendidos_com_data = (df_com_data['Número de atividades de vendas'] > 0).sum()
pct_atendidos_com_data = atendidos_com_data / len(df_com_data) * 100 if len(df_com_data) > 0 else 0

print(f"\nDos leads COM data válida:")
print(f"  - {atendidos_com_data:,} foram atendidos ({pct_atendidos_com_data:.1f}%)")
print(f"  - {len(df_com_data) - atendidos_com_data:,} NÃO foram atendidos ({100-pct_atendidos_com_data:.1f}%)")

# Nos leads SEM data válida, quantos foram atendidos?
atendidos_sem_data = (df_sem_data['Número de atividades de vendas'] > 0).sum()
pct_atendidos_sem_data = atendidos_sem_data / len(df_sem_data) * 100 if len(df_sem_data) > 0 else 0

print(f"\nDos leads SEM data válida:")
print(f"  - {atendidos_sem_data:,} foram atendidos ({pct_atendidos_sem_data:.1f}%)")
print(f"  - {len(df_sem_data) - atendidos_sem_data:,} NÃO foram atendidos ({100-pct_atendidos_sem_data:.1f}%)")

print("\n" + "="*80)
print("CONCLUSÃO:")
print("="*80)
print(f"Ao filtrar apenas leads com 'Data da última atividade' válida,")
print(f"estamos mantendo {pct_atendidos_com_data:.1f}% de leads atendidos (vs {pct_atendidos_sem_data:.1f}% na base descartada)")
print(f"Isso explica por que a aba FILTRADA tem percentuais mais altos!")
