import pandas as pd
from datetime import datetime

excel_30 = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-01-30.xlsx'
excel_06 = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-02-06.xlsx'
backup_path = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads_backup.csv'
atual_path = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\hubspot_leads.csv'

print("╔" + "═" * 140 + "╗")
print("║" + "ANÁLISE CRÍTICA: Excel → CSV → HubSpot Dashboard".center(140) + "║")
print("╚" + "═" * 140 + "╝\n")

# Carregar Excels
excel_30_df = pd.read_excel(excel_30, sheet_name=0)
excel_06_df = pd.read_excel(excel_06, sheet_name=0)

# Carregar CSVs
backup_df = pd.read_csv(backup_path, encoding='utf-8')
atual_df = pd.read_csv(atual_path, encoding='utf-8')
backup_df['Data de criação'] = pd.to_datetime(backup_df['Data de criação'], format='mixed', dayfirst=False)
atual_df['Data de criação'] = pd.to_datetime(atual_df['Data de criação'], format='mixed', dayfirst=False)

# Dados do HubSpot que o usuário vê HOJE
hubspot_hoje = {
    '23.1': 6225,
    '24.1': 8206,
    '25.1': 9698,
    '26.1': 10578,
}

# Ciclos
ciclos = {
    '23.1': (pd.Timestamp('2022-10-01'), pd.Timestamp('2023-01-31 23:59:59')),
    '24.1': (pd.Timestamp('2023-10-01'), pd.Timestamp('2024-01-31 23:59:59')),
    '25.1': (pd.Timestamp('2024-10-01'), pd.Timestamp('2025-01-31 23:59:59')),
    '26.1': (pd.Timestamp('2025-10-01'), pd.Timestamp('2026-01-31 23:59:59')),
}

print("TABELA COMPARATIVA - TODOS OS NÚMEROS\n")
print("Ciclo │ Excel 30/01 │ Excel 06/02 │ CSV Backup │ CSV Atual │ HubSpot Hoje │ Status")
print("      │  (Referência)│  (Base Atual)│ (30/01)    │  (06/02)  │  (Filtro UI) │")
print("─" * 140)

for ciclo_nome in ['23.1', '24.1', '25.1', '26.1']:
    excel_30_num = excel_30_df[excel_30_df['Ciclo'] == float(ciclo_nome)]['Leads'].values[0]
    excel_06_num = excel_06_df[excel_06_df['Ciclo'] == float(ciclo_nome)]['Leads'].values[0]
    
    data_inicio, data_fim = ciclos[ciclo_nome]
    backup_ciclo = backup_df[(backup_df['Data de criação'] >= data_inicio) & 
                              (backup_df['Data de criação'] <= data_fim)]
    atual_ciclo = atual_df[(atual_df['Data de criação'] >= data_inicio) & 
                            (atual_df['Data de criação'] <= data_fim)]
    
    csv_backup_num = len(backup_ciclo)
    csv_atual_num = len(atual_ciclo)
    hubspot_num = hubspot_hoje[ciclo_nome]
    
    # Determinar status
    status = ""
    if excel_30_num != csv_backup_num:
        status += f"⚠️ Excel30≠CSVBackup(-{csv_backup_num - excel_30_num}) "
    if excel_06_num != csv_atual_num:
        status += f"⚠️ Excel06≠CSVAtual(-{csv_atual_num - excel_06_num}) "
    if hubspot_num != csv_atual_num:
        status += f"❌ HubSpot≠CSV({hubspot_num - csv_atual_num:+d}) "
    if not status:
        status = "✓"
    
    print(f"{ciclo_nome:<4} │ {excel_30_num:>11,} │ {excel_06_num:>11,} │ {csv_backup_num:>10,} │ {csv_atual_num:>9,} │ {hubspot_num:>12,} │ {status}")

print("\n" + "═" * 140)
print("ANÁLISE DETALHADA POR CICLO")
print("═" * 140 + "\n")

for ciclo_nome in ['23.1', '24.1', '25.1', '26.1']:
    print(f"\n🔍 CICLO {ciclo_nome}")
    print("─" * 140 + "\n")
    
    excel_30_num = excel_30_df[excel_30_df['Ciclo'] == float(ciclo_nome)]['Leads'].values[0]
    excel_06_num = excel_06_df[excel_06_df['Ciclo'] == float(ciclo_nome)]['Leads'].values[0]
    
    data_inicio, data_fim = ciclos[ciclo_nome]
    backup_ciclo = backup_df[(backup_df['Data de criação'] >= data_inicio) & 
                              (backup_df['Data de criação'] <= data_fim)]
    atual_ciclo = atual_df[(atual_df['Data de criação'] >= data_inicio) & 
                            (atual_df['Data de criação'] <= data_fim)]
    
    csv_backup_num = len(backup_ciclo)
    csv_atual_num = len(atual_ciclo)
    hubspot_num = hubspot_hoje[ciclo_nome]
    
    print(f"  Excel 30/01 (reportado para Backup): {excel_30_num:,} leads")
    print(f"  CSV Backup real (30/01):             {csv_backup_num:,} leads")
    if excel_30_num != csv_backup_num:
        print(f"    → DIFERENÇA: {csv_backup_num - excel_30_num:+,} ({((csv_backup_num/excel_30_num - 1)*100):+.1f}%)\n")
    
    print(f"  Excel 06/02 (reportado para Atual):  {excel_06_num:,} leads")
    print(f"  CSV Atual real (06/02):              {csv_atual_num:,} leads")
    if excel_06_num != csv_atual_num:
        print(f"    → DIFERENÇA: {csv_atual_num - excel_06_num:+,} ({((csv_atual_num/excel_06_num - 1)*100):+.1f}%)\n")
    
    print(f"  HubSpot Dashboard HOJE (seu filtro): {hubspot_num:,} leads")
    print(f"  CSV Atual real (06/02):              {csv_atual_num:,} leads")
    if hubspot_num != csv_atual_num:
        print(f"    → DIFERENÇA: {hubspot_num - csv_atual_num:+,} ({((hubspot_num/csv_atual_num - 1)*100):+.1f}%)\n")
    
    # Investigar a discrepância
    print(f"  📊 Investigação da discrepância:")
    
    # Contar registros com data fora do range
    fora_do_range = atual_ciclo[
        (atual_ciclo['Data de criação'] < data_inicio) | 
        (atual_ciclo['Data de criação'] > data_fim)
    ]
    
    if len(fora_do_range) > 0:
        print(f"     • {len(fora_do_range)} registros DENTRO da seleção mas com data FORA do range")
    
    # Comparar com números do Excel
    if hubspot_num < csv_atual_num:
        diff = csv_atual_num - hubspot_num
        pct = (diff / csv_atual_num) * 100
        print(f"     • HubSpot mostra {diff:,} MENOS ({pct:.1f}%)")
        print(f"       → Possíveis razões:")
        print(f"         1. Filtros adicionais no Dashboard (pipeline, status, etc)")
        print(f"         2. Registros marcados como deletados não aparecem")
        print(f"         3. Registros com etapa diferente ou proprietário filtrado")

print("\n\n" + "═" * 140)
print("CONCLUSÃO E RECOMENDAÇÃO")
print("═" * 140 + "\n")

print("❓ O PROBLEMA:")
print("   • Excel 30/01 NÃO bate com CSV Backup (números diferentes)")
print("   • Excel 06/02 NÃO bate com CSV Atual (números diferentes)")
print("   • HubSpot HOJE NÃO bate com CSV Atual (números muito diferentes)")
print("\n🤔 POSSÍVEIS EXPLICAÇÕES:")
print("   1. Excel está calculando APENAS registros ativos (deletando deletados)")
print("   2. Excel tem filtros adicionais (etapa, pipeline, etc)")
print("   3. CSV inclui registros que deveriam estar excluídos")
print("   4. HubSpot Dashboard usa filtros diferentes da exportação CSV")
print("\n✅ RECOMENDAÇÃO:")
print("   • Verifique quais FILTROS o Excel aplica sobre os dados brutos")
print("   • Verifique os filtros do Dashboard do HubSpot que você utilizou")
print("   • Compare contagem de registros DELETADOS vs ATIVOS")
print("   • Valide se há registros com etapa específica que o Excel exclui")
