import pandas as pd
import numpy as np
from datetime import datetime
import os

print("=" * 140)
print("ANÁLISE FORENSE: DETECÇÃO DE MANIPULAÇÃO DE DADOS HUBSPOT")
print("Comparação: Base Backup (30/01/2026) vs Base Atual (06/02/2026)")
print("=" * 140)
print()

# Caminhos
base_backup = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads.csv'
base_atual = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\hubspot_leads.csv'

# Carregar ambas as bases
print("Carregando bases... (30/01/2026 e 06/02/2026)")
df_backup = pd.read_csv(base_backup)
df_atual = pd.read_csv(base_atual)

# Converter datas
df_backup['Data de criação'] = pd.to_datetime(df_backup['Data de criação'], format='%Y-%m-%d %H:%M', errors='coerce')
df_backup['Data de fechamento'] = pd.to_datetime(df_backup['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')

df_atual['Data de criação'] = pd.to_datetime(df_atual['Data de criação'], format='%Y-%m-%d %H:%M', errors='coerce')
df_atual['Data de fechamento'] = pd.to_datetime(df_atual['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')

print(f"✓ Base backup: {len(df_backup)} registros")
print(f"✓ Base atual: {len(df_atual)} registros")
print()

# ============================================================================
# 1. REGISTROS DELETADOS (estavam no backup, não estão na atual)
# ============================================================================
print("=" * 140)
print("1. ANÁLISE DE REGISTROS DELETADOS")
print("=" * 140)
print()

record_ids_backup = set(df_backup['Record ID'].astype(str))
record_ids_atual = set(df_atual['Record ID'].astype(str))

ids_deletados = record_ids_backup - record_ids_atual
ids_novos = record_ids_atual - record_ids_backup

print(f"Registros deletados (30/01 → 06/02): {len(ids_deletados)}")

if len(ids_deletados) > 0:
    print(f"\n🔴 EVIDÊNCIA CRÍTICA: {len(ids_deletados)} registros foram DELETADOS!")
    print("\nÚltimos 20 IDs deletados:")
    for record_id in sorted(ids_deletados)[-20:]:
        df_temp = df_backup[df_backup['Record ID'].astype(str) == record_id].iloc[0]
        print(f"  • ID: {record_id}")
        print(f"    Nome: {df_temp['Nome do negócio']}")
        print(f"    Criação: {df_temp['Data de criação']}")
        print(f"    Etapa: {df_temp['Etapa do negócio']}")
        print()

print(f"\nRegistros novos (30/01 → 06/02): {len(ids_novos)}")

# ============================================================================
# 2. REGISTROS MODIFICADOS
# ============================================================================
print("\n" + "=" * 140)
print("2. ANÁLISE DE REGISTROS MODIFICADOS")
print("=" * 140)
print()

ids_comuns = record_ids_backup & record_ids_atual
print(f"Registros em comum: {len(ids_comuns)}")
print()

# Comparar registros comuns
df_backup_comuns = df_backup[df_backup['Record ID'].astype(str).isin(ids_comuns)].set_index('Record ID')
df_atual_comuns = df_atual[df_atual['Record ID'].astype(str).isin(ids_comuns)].set_index('Record ID')

colunas_chave = ['Nome do negócio', 'Etapa do negócio', 'Data de criação', 'Data de fechamento', 
                  'Proprietário do negócio', 'Pipeline', 'Valor na moeda da empresa']

modificacoes = {}
registros_modificados = set()

for col in df_backup_comuns.columns:
    # Comparar valores
    diferentes = df_backup_comuns[col] != df_atual_comuns[col]
    
    # Lidar com NaN
    mask_nan_backup = df_backup_comuns[col].isna()
    mask_nan_atual = df_atual_comuns[col].isna()
    diferentes = diferentes & ~(mask_nan_backup & mask_nan_atual)
    
    qtd_diferentes = diferentes.sum()
    if qtd_diferentes > 0:
        modificacoes[col] = qtd_diferentes
        registros_modificados.update(diferentes[diferentes].index.astype(str))

print(f"Total de registros modificados: {len(registros_modificados)}")
print()

if modificacoes:
    print("Campos modificados (TOP 15):")
    for col, qtd in sorted(modificacoes.items(), key=lambda x: x[1], reverse=True)[:15]:
        print(f"  • {col}: {qtd} registros")
    print()
    
    # Mostrar 10 casos de modificação
    print("EXEMPLOS DE MODIFICAÇÕES SUSPEITAS:")
    print("-" * 140)
    
    for i, record_id in enumerate(sorted(registros_modificados)[:10]):
        print(f"\nRegistro #{i+1}: ID={record_id}")
        
        for col in df_backup_comuns.columns:
            val_backup = df_backup_comuns.loc[int(record_id), col]
            val_atual = df_atual_comuns.loc[int(record_id), col]
            
            # Comparar
            if pd.isna(val_backup) and pd.isna(val_atual):
                continue
            elif str(val_backup) != str(val_atual):
                print(f"  {col}:")
                print(f"    30/01 (Backup):   {val_backup}")
                print(f"    06/02 (Atual):    {val_atual}")

print()

# ============================================================================
# 3. ANÁLISE ESPECÍFICA DO PERÍODO OUTUBRO 24 - FEVEREIRO 25
# ============================================================================
print("\n" + "=" * 140)
print("3. ANÁLISE PERÍODO: OUTUBRO/24 - FEVEREIRO/25 (CICLO 26.1)")
print("=" * 140)
print()

data_inicio = pd.Timestamp('2024-10-01')
data_fim = pd.Timestamp('2025-02-28')

df_backup_periodo = df_backup[(df_backup['Data de criação'] >= data_inicio) & 
                               (df_backup['Data de criação'] <= data_fim)].copy()
df_atual_periodo = df_atual[(df_atual['Data de criação'] >= data_inicio) & 
                             (df_atual['Data de criação'] <= data_fim)].copy()

ids_backup_periodo = set(df_backup_periodo['Record ID'].astype(str))
ids_atual_periodo = set(df_atual_periodo['Record ID'].astype(str))

ids_deletados_periodo = ids_backup_periodo - ids_atual_periodo
ids_novos_periodo = ids_atual_periodo - ids_backup_periodo

print(f"Registros do período no backup: {len(df_backup_periodo)}")
print(f"Registros do período na atual: {len(df_atual_periodo)}")
print()

print(f"🔴 Registros DELETADOS neste período: {len(ids_deletados_periodo)}")
if len(ids_deletados_periodo) > 0:
    print("   IDs deletados (amostra de 10):")
    for record_id in list(ids_deletados_periodo)[:10]:
        df_temp = df_backup[df_backup['Record ID'].astype(str) == record_id].iloc[0]
        print(f"     • {record_id} - {df_temp['Nome do negócio']} - {df_temp['Etapa do negócio']}")

print(f"\n✓ Registros NOVOS neste período: {len(ids_novos_periodo)}")

# ============================================================================
# 4. ANÁLISE DE MATRÍCULAS DELETADAS
# ============================================================================
print("\n" + "=" * 140)
print("4. ANÁLISE DE MATRÍCULAS DELETADAS")
print("=" * 140)
print()

# Matrículas no backup
df_mat_backup = df_backup[df_backup['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)]
# Matrículas na atual
df_mat_atual = df_atual[df_atual['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)]

ids_mat_backup = set(df_mat_backup['Record ID'].astype(str))
ids_mat_atual = set(df_mat_atual['Record ID'].astype(str))

ids_mat_deletadas = ids_mat_backup - ids_mat_atual

print(f"Matrículas no backup (30/01): {len(df_mat_backup)}")
print(f"Matrículas na atual (06/02): {len(df_mat_atual)}")
print(f"\n🔴 MATRÍCULAS DELETADAS: {len(ids_mat_deletadas)}")

if len(ids_mat_deletadas) > 0:
    print("\n⚠️  EVIDÊNCIA CRÍTICA: Matrículas foram REMOVIDAS do banco de dados!")
    print("\nExemplos de matrículas deletadas:")
    for record_id in list(ids_mat_deletadas)[:10]:
        df_temp = df_mat_backup[df_mat_backup['Record ID'].astype(str) == record_id].iloc[0]
        print(f"  • {record_id} - {df_temp['Nome do negócio']}")
        print(f"    Cliente: {df_temp['Proprietário do negócio']}")
        print(f"    Fechamento: {df_temp['Data de fechamento']}")
        print()

# ============================================================================
# RESUMO FINAL
# ============================================================================
print("\n" + "=" * 140)
print("RESUMO EXECUTIVO - EVIDÊNCIAS DE MANIPULAÇÃO")
print("=" * 140)
print()

print(f"📊 ESTATÍSTICAS GERAIS:")
print(f"   • Registros deletados (geral): {len(ids_deletados)}")
print(f"   • Registros modificados: {len(registros_modificados)}")
print(f"   • Registros novos: {len(ids_novos)}")
print()

print(f"📊 PERÍODO OUTUBRO/24 - FEVEREIRO/25:")
print(f"   • Registros deletados neste período: {len(ids_deletados_periodo)}")
print(f"   • Registros novos neste período: {len(ids_novos_periodo)}")
print()

print(f"📊 MATRÍCULAS:")
print(f"   • Matrículas deletadas (CRÍTICO): {len(ids_mat_deletadas)}")
print()

# Exportar relatório
output_file = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\relatorio_auditoria_dados\ANALISE_FORENSE_MANIPULACAO.txt'
os.makedirs(os.path.dirname(output_file), exist_ok=True)

with open(output_file, 'w', encoding='utf-8') as f:
    f.write("ANÁLISE FORENSE: DETECÇÃO DE MANIPULAÇÃO HUBSPOT\n")
    f.write("=" * 140 + "\n")
    f.write(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
    f.write(f"Comparação: 30/01/2026 vs 06/02/2026\n\n")
    
    f.write("RESULTADO DA ANÁLISE:\n\n")
    f.write(f"Registros deletados: {len(ids_deletados)}\n")
    f.write(f"Registros modificados: {len(registros_modificados)}\n")
    f.write(f"Matrículas deletadas: {len(ids_mat_deletadas)}\n")

print(f"\n✓ Relatório salvo em: {output_file}")
