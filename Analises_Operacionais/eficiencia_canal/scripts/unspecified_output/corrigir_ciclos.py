import pandas as pd
from datetime import datetime

print("=" * 140)
print("RECALCULANDO COM PERÍODOS CORRETOS")
print("=" * 140)
print()

# Carregar bases
base_backup = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads_backup.csv'
base_atual = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\hubspot_leads.csv'

df_backup = pd.read_csv(base_backup)
df_atual = pd.read_csv(base_atual)

df_backup['Data de criação'] = pd.to_datetime(df_backup['Data de criação'], format='%Y-%m-%d %H:%M', errors='coerce')
df_atual['Data de criação'] = pd.to_datetime(df_atual['Data de criação'], format='%Y-%m-%d %H:%M', errors='coerce')

# Ciclos CORRETOS baseado no HubSpot e análise do dia 30
# Ciclo 25.1: Outubro 2024 a 30 Janeiro 2025
# Ciclo 26.1: Fevereiro 2025 em diante

print("CICLO 25.1: Outubro 2024 a 30 Janeiro 2025")
print("-" * 140)

# Periodo correto para 25.1
data_inicio_25_1 = pd.Timestamp('2024-10-01')
data_fim_25_1 = pd.Timestamp('2025-01-30 23:59:59')

df_25_1_backup = df_backup[(df_backup['Data de criação'] >= data_inicio_25_1) & 
                            (df_backup['Data de criação'] <= data_fim_25_1)]
df_25_1_atual = df_atual[(df_atual['Data de criação'] >= data_inicio_25_1) & 
                          (df_atual['Data de criação'] <= data_fim_25_1)]

print(f"Backup (30/01): {len(df_25_1_backup)} leads")
print(f"Atual (06/02): {len(df_25_1_atual)} leads")
print(f"Análise dia 30: 9.688 leads")
print(f"HubSpot hoje: 9.698 leads")
print()

# Agrupar por mês
print("Quebra por mês (Backup):")
df_25_1_backup['Mes'] = df_25_1_backup['Data de criação'].dt.to_period('M')
por_mes = df_25_1_backup.groupby('Mes').size()
for mes, qtd in por_mes.items():
    print(f"  {mes}: {qtd}")
print()

# IDs em comum
ids_backup_25_1 = set(df_25_1_backup['Record ID'].astype(str))
ids_atual_25_1 = set(df_25_1_atual['Record ID'].astype(str))

ids_deletados = ids_backup_25_1 - ids_atual_25_1
ids_novos = ids_atual_25_1 - ids_backup_25_1

print(f"Deletados neste período: {len(ids_deletados)}")
print(f"Novos neste período: {len(ids_novos)}")
print()

# Comparar os números
print("=" * 140)
print("DISCREPÂNCIA DE 10 LEADS")
print("=" * 140)
print()

print("Backup tem 10 leads a MENOS que o HubSpot.")
print("Possíveis causas:")
print("1. Registros duplicados no HubSpot que backup não tem?")
print("2. Registros sem 'Data de criação' no backup?")
print("3. Filtro de data diferente?")
print()

# Verificar registros sem data
sem_data_backup = df_backup[df_backup['Data de criação'].isna()]
sem_data_atual = df_atual[df_atual['Data de criação'].isna()]

print(f"Registros SEM data criaçao - Backup: {len(sem_data_backup)}")
print(f"Registros SEM data criaçao - Atual: {len(sem_data_atual)}")
print()

# Verificar se há duplicados
dups_backup = df_25_1_backup['Record ID'].duplicated().sum()
dups_atual = df_25_1_atual['Record ID'].duplicated().sum()

print(f"Record IDs duplicados - Backup: {dups_backup}")
print(f"Record IDs duplicados - Atual: {dups_atual}")
print()

# Matrículas do período
print("=" * 140)
print("MATRÍCULAS - CICLO 25.1")
print("=" * 140)
print()

df_mat_backup_25_1 = df_25_1_backup[df_25_1_backup['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)]
df_mat_atual_25_1 = df_25_1_atual[df_25_1_atual['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)]

print(f"Matrículas - Backup: {len(df_mat_backup_25_1)}")
print(f"Matrículas - Atual: {len(df_mat_atual_25_1)}")
print(f"Análise dia 30: 2.800 matrículas")
print()
