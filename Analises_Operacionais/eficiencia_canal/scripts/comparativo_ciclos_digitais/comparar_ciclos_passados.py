import pandas as pd
import os
from datetime import datetime

# Caminhos das bases
backup_path = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads_backup.csv'
atual_path = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\hubspot_leads.csv'

print("╔" + "═" * 120 + "╗")
print("║" + "COMPARAÇÃO: Ciclos Passados - Backup(30/01) vs Atual(06/02)".center(120) + "║")
print("╚" + "═" * 120 + "╝\n")

# Carregar dados
print("📂 Carregando bases...")
backup_df = pd.read_csv(backup_path, encoding='utf-8')
atual_df = pd.read_csv(atual_path, encoding='utf-8')

print(f"   Backup (30/01): {len(backup_df):,} registros")
print(f"   Atual (06/02):  {len(atual_df):,} registros")
print(f"   Diferença: {len(atual_df) - len(backup_df):+,} registros\n")

# Converter Data de criação
backup_df['Data de criação'] = pd.to_datetime(backup_df['Data de criação'], format='mixed', dayfirst=False)
atual_df['Data de criação'] = pd.to_datetime(atual_df['Data de criação'], format='mixed', dayfirst=False)

# Definir ciclos corretos (OUT-JAN)
ciclos = {
    '23.1': (pd.Timestamp('2022-10-01'), pd.Timestamp('2023-01-31 23:59:59')),
    '24.1': (pd.Timestamp('2023-10-01'), pd.Timestamp('2024-01-31 23:59:59')),
    '25.1': (pd.Timestamp('2024-10-01'), pd.Timestamp('2025-01-31 23:59:59')),
}

print("─" * 120)
print("CICLOS PASSADOS - Análise")
print("─" * 120 + "\n")

for ciclo_nome, (data_inicio, data_fim) in ciclos.items():
    print(f"\n🔍 CICLO {ciclo_nome} ({data_inicio.strftime('%d/%m/%Y')} - {data_fim.strftime('%d/%m/%Y')})")
    print("   " + "─" * 116)
    
    # Filtrar ciclo em ambas bases
    backup_ciclo = backup_df[(backup_df['Data de criação'] >= data_inicio) & 
                             (backup_df['Data de criação'] <= data_fim)]
    atual_ciclo = atual_df[(atual_df['Data de criação'] >= data_inicio) & 
                           (atual_df['Data de criação'] <= data_fim)]
    
    print(f"   BACKUP (30/01): {len(backup_ciclo):,} registros")
    print(f"   ATUAL (06/02):  {len(atual_ciclo):,} registros")
    
    # Comparar IDs
    ids_backup = set(backup_ciclo['Record ID'].astype(str))
    ids_atual = set(atual_ciclo['Record ID'].astype(str))
    
    deletados = ids_backup - ids_atual
    novos = ids_atual - ids_backup
    comuns = ids_backup & ids_atual
    
    print(f"\n   Resultado:")
    print(f"   • Registros em comum: {len(comuns):,}")
    print(f"   • Registros DELETADOS: {len(deletados)}")
    print(f"   • Registros ADICIONADOS: {len(novos)}")
    
    if len(deletados) > 0:
        print(f"\n   🔴 ALERTA: {len(deletados)} registros foram DELETADOS deste ciclo!")
        print(f"      IDs: {', '.join(sorted(list(deletados))[:10])}")
        if len(deletados) > 10:
            print(f"      ... e mais {len(deletados) - 10}")
    
    if len(novos) > 0:
        print(f"\n   🟡 ALERTA: {len(novos)} registros foram ADICIONADOS neste ciclo!")
        print(f"      IDs: {', '.join(sorted(list(novos))[:10])}")
        if len(novos) > 10:
            print(f"      ... e mais {len(novos) - 10}")
    
    # Se houver registros em comum, verificar mudanças
    if len(comuns) > 0:
        backup_comuns = backup_ciclo[backup_ciclo['Record ID'].astype(str).isin(comuns)].set_index('Record ID')
        atual_comuns = atual_ciclo[atual_ciclo['Record ID'].astype(str).isin(comuns)].set_index('Record ID')
        
        # Colunas críticas para checar
        colunas_criticas = ['Etapa do negócio', 'Proprietário do negócio', 'Data de fechamento']
        
        mudancas_total = 0
        for coluna in colunas_criticas:
            if coluna in backup_comuns.columns and coluna in atual_comuns.columns:
                # Comparar valores
                mask = backup_comuns[coluna].astype(str) != atual_comuns[coluna].astype(str)
                mudancas = mask.sum()
                mudancas_total += mudancas
                
                if mudancas > 0:
                    print(f"\n   🔴 {coluna}: {mudancas} registros MODIFICADOS")
                    # Mostrar exemplos
                    exemplos = backup_comuns[mask].head(3)
                    for idx in exemplos.index[:3]:
                        antes = backup_comuns.loc[idx, coluna]
                        depois = atual_comuns.loc[idx, coluna]
                        print(f"      ID {idx}: '{antes}' → '{depois}'")
        
        if mudancas_total == 0:
            print(f"\n   ✓ NENHUMA mudança detectada nos registros em comum")

print("\n\n" + "═" * 120)
print("RESUMO EXECUTIVO - CICLOS PASSADOS")
print("═" * 120 + "\n")

print("⚠️ ANÁLISE CRÍTICA:\n")
print("Os ciclos PASSADOS (23.1, 24.1, 25.1) deveriam estar CONGELADOS.")
print("Qualquer mudança neles é indicativo de:")
print("  • Manipulação intencional de dados históricos")
print("  • Alteração retroativa de relatórios")
print("  • Comprometimento da integridade de dados\n")

print("# Comparação geral ciclos passados")
total_mudancas = 0
for ciclo_nome in ['23.1', '24.1', '25.1']:
    data_inicio, data_fim = ciclos[ciclo_nome]
    backup_ciclo = backup_df[(backup_df['Data de criação'] >= data_inicio) & 
                             (backup_df['Data de criação'] <= data_fim)]
    atual_ciclo = atual_df[(atual_df['Data de criação'] >= data_inicio) & 
                           (atual_df['Data de criação'] <= data_fim)]
    
    ids_backup = set(backup_ciclo['Record ID'].astype(str))
    ids_atual = set(atual_ciclo['Record ID'].astype(str))
    deletados = ids_backup - ids_atual
    novos = ids_atual - ids_backup
    
    total_mudancas += len(deletados) + len(novos)
    
    if len(deletados) > 0 or len(novos) > 0:
        print(f"❌ CICLO {ciclo_nome}: {len(deletados)} deletados, {len(novos)} novos")

if total_mudancas > 0:
    print(f"\n🔴 CONCLUSÃO: {total_mudancas} registros foram alterados em ciclos PASSADOS!")
    print("   Isso é GRAVE - estes ciclos deveriam estar imutáveis.")
else:
    print("\n✓ CONCLUSÃO: Ciclos passados estão INTACTOS")
    print("  Os dados do dia 30/01 refletem corretamente a base atual para ciclos 23.1, 24.1, 25.1")

print("\n" + "═" * 120)
