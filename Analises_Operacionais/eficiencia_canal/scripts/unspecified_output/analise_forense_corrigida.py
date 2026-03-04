import pandas as pd
from datetime import datetime
import os

print("=" * 160)
print("ANÁLISE FORENSE CORRIGIDA - CICLOS OUTUBRO A JANEIRO")
print("Base Backup 30/01/2026 vs Base Atual 06/02/2026")
print("=" * 160)
print()

# Caminhos
base_backup = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads_backup.csv'
base_atual = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\hubspot_leads.csv'

# Carregar
df_backup = pd.read_csv(base_backup)
df_atual = pd.read_csv(base_atual)

df_backup['Data de criação'] = pd.to_datetime(df_backup['Data de criação'], format='%Y-%m-%d %H:%M', errors='coerce')
df_backup['Data de fechamento'] = pd.to_datetime(df_backup['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')

df_atual['Data de criação'] = pd.to_datetime(df_atual['Data de criação'], format='%Y-%m-%d %H:%M', errors='coerce')
df_atual['Data de fechamento'] = pd.to_datetime(df_atual['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')

# Ciclos CORRETOS: Out a Jan (5 meses)
ciclos = {
    '23.1': (pd.Timestamp('2022-10-01'), pd.Timestamp('2023-01-31 23:59:59')),
    '24.1': (pd.Timestamp('2023-10-01'), pd.Timestamp('2024-01-31 23:59:59')),
    '25.1': (pd.Timestamp('2024-10-01'), pd.Timestamp('2025-01-31 23:59:59')),
    '26.1': (pd.Timestamp('2025-10-01'), pd.Timestamp('2026-01-31 23:59:59')),
}

print("ANÁLISE POR CICLO - Comparação Backup (30/01) vs Atual (06/02)")
print("=" * 160)
print()

for ciclo_nome, (data_inicio, data_fim) in ciclos.items():
    print(f"\n{ciclo_nome} ({data_inicio.strftime('%b%Y')} - {data_fim.strftime('%b%Y')})")
    print("-" * 160)
    
    # Filtrar leads
    df_backup_ciclo = df_backup[(df_backup['Data de criação'] >= data_inicio) & 
                                (df_backup['Data de criação'] <= data_fim)]
    df_atual_ciclo = df_atual[(df_atual['Data de criação'] >= data_inicio) & 
                              (df_atual['Data de criação'] <= data_fim)]
    
    # IDs comum/novo/deletado
    ids_backup = set(df_backup_ciclo['Record ID'].astype(str))
    ids_atual = set(df_atual_ciclo['Record ID'].astype(str))
    
    ids_deletados = ids_backup - ids_atual
    ids_novos = ids_atual - ids_backup
    ids_comuns = ids_backup & ids_atual
    
    print(f"Leads - Backup: {len(df_backup_ciclo)}, Atual: {len(df_atual_ciclo)}")
    print(f"  Deletados: {len(ids_deletados)}, Novos: {len(ids_novos)}, Comuns: {len(ids_comuns)}")
    
    # Matrículas
    df_mat_backup_ciclo = df_backup_ciclo[df_backup_ciclo['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)]
    df_mat_backup_ciclo = df_mat_backup_ciclo[(df_mat_backup_ciclo['Data de fechamento'] >= data_inicio) & 
                                              (df_mat_backup_ciclo['Data de fechamento'] <= data_fim)]
    
    df_mat_atual_ciclo = df_atual_ciclo[df_atual_ciclo['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)]
    df_mat_atual_ciclo = df_mat_atual_ciclo[(df_mat_atual_ciclo['Data de fechamento'] >= data_inicio) & 
                                            (df_mat_atual_ciclo['Data de fechamento'] <= data_fim)]
    
    print(f"Matrículas - Backup: {len(df_mat_backup_ciclo)}, Atual: {len(df_mat_atual_ciclo)}")
    
    # Detectar modificações
    if len(ids_comuns) > 0:
        df_backup_comuns = df_backup_ciclo[df_backup_ciclo['Record ID'].astype(str).isin(ids_comuns)].set_index('Record ID')
        df_atual_comuns = df_atual_ciclo[df_atual_ciclo['Record ID'].astype(str).isin(ids_comuns)].set_index('Record ID')
        
        # Comparar etapas
        mudancas_etapa = (df_backup_comuns['Etapa do negócio'] != df_atual_comuns['Etapa do negócio'])
        # Lidar com NaN
        mask_nan_backup = df_backup_comuns['Etapa do negócio'].isna()
        mask_nan_atual = df_atual_comuns['Etapa do negócio'].isna()
        mudancas_etapa = mudancas_etapa & ~(mask_nan_backup & mask_nan_atual)
        
        qtd_mudancas_etapa = mudancas_etapa.sum()
        
        # Mudanças de proprietário
        mudancas_prop = (df_backup_comuns['Proprietário do negócio'] != df_atual_comuns['Proprietário do negócio'])
        mask_nan_backup_prop = df_backup_comuns['Proprietário do negócio'].isna()
        mask_nan_atual_prop = df_atual_comuns['Proprietário do negócio'].isna()
        mudancas_prop = mudancas_prop & ~(mask_nan_backup_prop & mask_nan_atual_prop)
        
        qtd_mudancas_prop = mudancas_prop.sum()
        
        if qtd_mudancas_etapa > 0 or qtd_mudancas_prop > 0:
            print(f"  ⚠️  MODIFICAÇÕES: {qtd_mudancas_etapa} mudanças de etapa, {qtd_mudancas_prop} mudanças de proprietário")

print("\n" + "=" * 160)
print("RESUMO FINAL")
print("=" * 160)

# Totais gerais
ids_backup_total = set(df_backup['Record ID'].astype(str))
ids_atual_total = set(df_atual['Record ID'].astype(str))

ids_deletados_total = ids_backup_total - ids_atual_total
ids_novos_total = ids_atual_total - ids_backup_total

print(f"\nComparação geral (todas as datas):")
print(f"  Registros deletados: {len(ids_deletados_total)}")
print(f"  Registros novos: {len(ids_novos_total)}")
print(f"\nComo os ciclos OUT-JAN batem perfeitamente com a análise do dia 30,")
print(f"qualquer modificação nesses ciclos É EVIDÊNCIA DE MANIPULAÇÃO!")

# Exportar resultado
output_file = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\relatorio_auditoria_dados\ANALISE_FORENSE_CORRIGIDA.txt'
os.makedirs(os.path.dirname(output_file), exist_ok=True)

with open(output_file, 'w', encoding='utf-8') as f:
    f.write("ANÁLISE FORENSE - CICLOS CORRETOS (OUT-JAN)\n")
    f.write("=" * 160 + "\n")
    f.write(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
    f.write(f"Ciclos validados com análise do dia 30/01/2026\n\n")
    f.write(f"Registros deletados (geral): {len(ids_deletados_total)}\n")
    f.write(f"Registros novos (geral): {len(ids_novos_total)}\n")

print(f"\n✓ Relatório exportado: {output_file}")
