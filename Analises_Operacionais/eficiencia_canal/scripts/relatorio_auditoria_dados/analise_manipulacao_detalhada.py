import pandas as pd
import numpy as np
from datetime import datetime
import os

print("=" * 140)
print("ANÁLISE DETALHADA DE MANIPULAÇÃO - FOCO NAS MUDANÇAS SUSPEITAS")
print("=" * 140)
print()

# Caminhos
base_backup = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads.csv'
base_atual = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\hubspot_leads.csv'

# Carregar
df_backup = pd.read_csv(base_backup)
df_atual = pd.read_csv(base_atual)

df_backup['Data de criação'] = pd.to_datetime(df_backup['Data de criação'], format='%Y-%m-%d %H:%M', errors='coerce')
df_backup['Data de fechamento'] = pd.to_datetime(df_backup['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')

df_atual['Data de criação'] = pd.to_datetime(df_atual['Data de criação'], format='%Y-%m-%d %H:%M', errors='coerce')
df_atual['Data de fechamento'] = pd.to_datetime(df_atual['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')

# IDs comuns
record_ids_backup = set(df_backup['Record ID'].astype(str))
record_ids_atual = set(df_atual['Record ID'].astype(str))
ids_comuns = record_ids_backup & record_ids_atual

df_backup_comuns = df_backup[df_backup['Record ID'].astype(str).isin(ids_comuns)].set_index('Record ID').sort_index()
df_atual_comuns = df_atual[df_atual['Record ID'].astype(str).isin(ids_comuns)].set_index('Record ID').sort_index()

# ============================================================================
# 1. MUDANÇAS DE ETAPA SUSPEITAS
# ============================================================================
print("=" * 140)
print("1. MUDANÇAS DE ETAPA DO NEGÓCIO (POTENCIAL MANIPULAÇÃO)")
print("=" * 140)
print()

mudancas_etapa = (df_backup_comuns['Etapa do negócio'] != df_atual_comuns['Etapa do negócio'])
mudancas_etapa = mudancas_etapa & ~(df_backup_comuns['Etapa do negócio'].isna() & df_atual_comuns['Etapa do negócio'].isna())
ids_mudancas_etapa = mudancas_etapa[mudancas_etapa].index.astype(str)

print(f"Total de registros com mudança de etapa: {len(ids_mudancas_etapa)}\n")

# Categorizar mudanças
categorias_mudancas = {}
for record_id in ids_mudancas_etapa[:50]:
    etapa_antes = df_backup_comuns.loc[int(record_id), 'Etapa do negócio']
    etapa_depois = df_atual_comuns.loc[int(record_id), 'Etapa do negócio']
    
    chave = f"{etapa_antes} → {etapa_depois}"
    if chave not in categorias_mudancas:
        categorias_mudancas[chave] = []
    categorias_mudancas[chave].append(record_id)

print("Tipos de mudanças de etapa (TOP 10):")
for mudanca, ids in sorted(categorias_mudancas.items(), key=lambda x: len(x[1]), reverse=True)[:10]:
    print(f"  {mudanca}: {len(ids)} registros")
print()

# Mostrar exemplos
print("EXEMPLOS SUSPEITOS - Registros que mudaram de etapa:")
print("-" * 140)

for record_id in list(ids_mudancas_etapa)[:5]:
    etapa_antes = df_backup_comuns.loc[int(record_id), 'Etapa do negócio']
    etapa_depois = df_atual_comuns.loc[int(record_id), 'Etapa do negócio']
    nome = df_backup_comuns.loc[int(record_id), 'Nome do negócio']
    
    print(f"\n• Record ID: {record_id}")
    print(f"  Nome: {nome}")
    print(f"  Mudança: {etapa_antes} → {etapa_depois}")
    
    # Se era matrícula concluída e agora não é, é CRÍTICO
    if "MATRÍCULA CONCLUÍDA" in str(etapa_antes) and "MATRÍCULA CONCLUÍDA" not in str(etapa_depois):
        print(f"  🔴 ALERTA: Matrícula foi DESCLASSIFICADA!")

# ============================================================================
# 2. MUDANÇAS DE PROPRIETÁRIO SUSPEITAS
# ============================================================================
print("\n\n" + "=" * 140)
print("2. MUDANÇAS DE PROPRIETÁRIO (REASSIGN SUSPEITO)")
print("=" * 140)
print()

mudancas_prop = (df_backup_comuns['Proprietário do negócio'] != df_atual_comuns['Proprietário do negócio'])
mudancas_prop = mudancas_prop & ~(df_backup_comuns['Proprietário do negócio'].isna() & df_atual_comuns['Proprietário do negócio'].isna())
ids_mudancas_prop = mudancas_prop[mudancas_prop].index.astype(str)

print(f"Total de registros com mudança de proprietário: {len(ids_mudancas_prop)}\n")

# Mostrar principais mudanças
prop_mudancas = {}
for record_id in ids_mudancas_prop:
    prop_antes = df_backup_comuns.loc[int(record_id), 'Proprietário do negócio']
    prop_depois = df_atual_comuns.loc[int(record_id), 'Proprietário do negócio']
    
    chave = f"{prop_antes} → {prop_depois}"
    if chave not in prop_mudancas:
        prop_mudancas[chave] = 0
    prop_mudancas[chave] += 1

print("Principais reassigns de proprietário (TOP 5):")
for mudanca, qtd in sorted(prop_mudancas.items(), key=lambda x: x[1], reverse=True)[:5]:
    print(f"  {mudanca}: {qtd} registros")

# ============================================================================
# 3. MATRÍCULAS COM ETAPA MODIFICADA
# ============================================================================
print("\n\n" + "=" * 140)
print("3. MATRÍCULAS COM ETAPA MODIFICADA (CRÍTICO)")
print("=" * 140)
print()

# Matrículas no backup que foram modificadas
df_mat_backup = df_backup_comuns[df_backup_comuns['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)]

# Ver quais tiveram etapa modificada
mat_etapa_mudada = mudancas_etapa[mudancas_etapa].index.intersection(df_mat_backup.index)

print(f"Matrículas com etapa modificada: {len(mat_etapa_mudada)}\n")

if len(mat_etapa_mudada) > 0:
    print("🔴 EVIDÊNCIA CRÍTICA: Matrículas tiveram suas etapas alteradas!\n")
    for record_id in mat_etapa_mudada:
        etapa_antes = df_backup_comuns.loc[record_id, 'Etapa do negócio']
        etapa_depois = df_atual_comuns.loc[record_id, 'Etapa do negócio']
        nome = df_backup_comuns.loc[record_id, 'Nome do negócio']
        data_fech = df_backup_comuns.loc[record_id, 'Data de fechamento']
        
        print(f"  • {record_id} - {nome}")
        print(f"    De: {etapa_antes} → Para: {etapa_depois}")
        print(f"    Data fechamento: {data_fech}\n")

# ============================================================================
# 4. REGISTROS COM DATA DE ATIVIDADE ALTERADA SIGNIFICANTEMENTE
# ============================================================================
print("\n" + "=" * 140)
print("4. REGISTROS COM DATA DE ATIVIDADE ALTERADA (TIMELINE MODIFICADA)")
print("=" * 140)
print()

data_ativ_mudada = (df_backup_comuns['Data da última atividade'] != df_atual_comuns['Data da última atividade'])
data_ativ_mudada = data_ativ_mudada & ~(df_backup_comuns['Data da última atividade'].isna() & df_atual_comuns['Data da última atividade'].isna())
ids_data_mudada = data_ativ_mudada[data_ativ_mudada].index.astype(str)

print(f"Total de registros com data de atividade alterada: {len(ids_data_mudada)}\n")

# Ver diferenças significativas
print("Exemplos de alterações na data de última atividade (diferenças > 1 mês):")
print("-" * 140)

diferenca_grande = 0
for record_id in list(ids_data_mudada)[:100]:
    data_antes = pd.to_datetime(df_backup_comuns.loc[int(record_id), 'Data da última atividade'], errors='coerce')
    data_depois = pd.to_datetime(df_atual_comuns.loc[int(record_id), 'Data da última atividade'], errors='coerce')
    
    if pd.notna(data_antes) and pd.notna(data_depois):
        diferenca = (data_depois - data_antes).days
        
        if abs(diferenca) > 30:  # Mais de 1 mês de diferença
            diferenca_grande += 1
            if diferenca_grande <= 5:
                nome = df_backup_comuns.loc[int(record_id), 'Nome do negócio']
                print(f"\n  • {record_id}")
                print(f"    Nome: {nome}")
                print(f"    30/01: {data_antes} → 06/02: {data_depois}")
                print(f"    Diferença: {diferenca} dias")

print(f"\nTotal com diferença > 30 dias: {diferenca_grande}")

# ============================================================================
# EXPORTAR RELATÓRIO DETALHADO
# ============================================================================
print("\n\n" + "=" * 140)
print("EXPORTANDO RELATÓRIO DETALHADO...")
print("=" * 140)

output_file = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\relatorio_auditoria_dados\MANIPULACAO_DETALHADA.txt'
os.makedirs(os.path.dirname(output_file), exist_ok=True)

with open(output_file, 'w', encoding='utf-8') as f:
    f.write("ANÁLISE DETALHADA DE MANIPULAÇÃO HUBSPOT\n")
    f.write("=" * 140 + "\n")
    f.write(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
    f.write(f"Período: 30/01/2026 (Backup) vs 06/02/2026 (Atual)\n\n")
    
    f.write("ACHADOS PRINCIPAIS:\n\n")
    
    f.write(f"1. Mudanças de Etapa: {len(ids_mudancas_etapa)} registros\n")
    f.write(f"2. Mudanças de Proprietário: {len(ids_mudancas_prop)} registros\n")
    f.write(f"3. Matrículas com Etapa Modificada: {len(mat_etapa_mudada)} registros (🔴 CRÍTICO)\n")
    f.write(f"4. Data de Atividade Alterada: {len(ids_data_mudada)} registros\n")
    f.write(f"   - Diferenças > 30 dias: {diferenca_grande} registros\n\n")
    
    f.write("CONCLUSÃO:\n")
    f.write("Os dados foram significativamente alterados entre 30/01 e 06/02.\n")
    f.write("Existem evidências de manipulação, especialmente em etapas de negócio e datas.\n")

print(f"✓ Relatório salvo em: {output_file}\n")
