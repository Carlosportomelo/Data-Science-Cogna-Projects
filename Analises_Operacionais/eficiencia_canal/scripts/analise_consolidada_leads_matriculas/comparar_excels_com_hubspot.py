import pandas as pd
import os
from datetime import datetime

print("╔" + "═" * 130 + "╗")
print("║" + "AUDITORIA CRÍTICA: Excel 30/01 vs Excel 06/02 vs HubSpot HOJE vs Bases CSV".center(130) + "║")
print("╚" + "═" * 130 + "╝\n")

# Caminhos
excel_30 = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-01-30.xlsx'
excel_06 = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-02-06.xlsx'
backup_path = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads_backup.csv'
atual_path = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\hubspot_leads.csv'

# Dados que o usuário vê HOJE no HubSpot
hubspot_hoje = {
    '23.1': 6225,
    '24.1': 8206,
    '25.1': 9698,
    '26.1': 10578,
}

print("📊 CARREGANDO ARQUIVOS...\n")

# Carregar Excels
try:
    excel_30_df = pd.read_excel(excel_30, sheet_name='Ciclos')
    print(f"✓ Excel 30/01 carregado: {len(excel_30_df)} linhas")
except Exception as e:
    print(f"✗ Erro ao carregar Excel 30/01: {e}")
    excel_30_df = None

try:
    excel_06_df = pd.read_excel(excel_06, sheet_name='Ciclos')
    print(f"✓ Excel 06/02 carregado: {len(excel_06_df)} linhas")
except Exception as e:
    print(f"✗ Erro ao carregar Excel 06/02: {e}")
    excel_06_df = None

# Carregar CSVs
backup_df = pd.read_csv(backup_path, encoding='utf-8')
atual_df = pd.read_csv(atual_path, encoding='utf-8')

# Converter datas
backup_df['Data de criação'] = pd.to_datetime(backup_df['Data de criação'], format='mixed', dayfirst=False)
atual_df['Data de criação'] = pd.to_datetime(atual_df['Data de criação'], format='mixed', dayfirst=False)

print(f"✓ CSV Backup carregado: {len(backup_df)} registros")
print(f"✓ CSV Atual carregado: {len(atual_df)} registros\n")

# Ciclos
ciclos = {
    '23.1': (pd.Timestamp('2022-10-01'), pd.Timestamp('2023-01-31 23:59:59')),
    '24.1': (pd.Timestamp('2023-10-01'), pd.Timestamp('2024-01-31 23:59:59')),
    '25.1': (pd.Timestamp('2024-10-01'), pd.Timestamp('2025-01-31 23:59:59')),
    '26.1': (pd.Timestamp('2025-10-01'), pd.Timestamp('2026-01-31 23:59:59')),
}

print("─" * 130)
print("COMPARAÇÃO POR CICLO - LEADS (sem matrículas)")
print("─" * 130 + "\n")

for ciclo_nome in ['23.1', '24.1', '25.1', '26.1']:
    data_inicio, data_fim = ciclos[ciclo_nome]
    
    print(f"\n{'═' * 130}")
    print(f"CICLO {ciclo_nome} ({data_inicio.strftime('%d/%m/%Y')} - {data_fim.strftime('%d/%m/%Y')})")
    print(f"{'═' * 130}")
    
    # Contar no CSV Backup
    backup_ciclo = backup_df[(backup_df['Data de criação'] >= data_inicio) & 
                              (backup_df['Data de criação'] <= data_fim)]
    count_backup = len(backup_ciclo)
    
    # Contar no CSV Atual
    atual_ciclo = atual_df[(atual_df['Data de criação'] >= data_inicio) & 
                            (atual_df['Data de criação'] <= data_fim)]
    count_atual = len(atual_ciclo)
    
    # Dados do Excel 30
    count_excel_30 = None
    if excel_30_df is not None:
        try:
            ciclo_row = excel_30_df[excel_30_df['Ciclo'] == float(ciclo_nome.split('.')[0]) + float(ciclo_nome.split('.')[1])/10]
            if not ciclo_row.empty:
                # Procurar coluna de leads
                leads_col = None
                for col in excel_30_df.columns:
                    if 'leads' in col.lower() and 'matriz' not in col.lower():
                        leads_col = col
                        break
                if leads_col:
                    count_excel_30 = ciclo_row[leads_col].values[0]
        except:
            pass
    
    # Dados do Excel 06
    count_excel_06 = None
    if excel_06_df is not None:
        try:
            ciclo_row = excel_06_df[excel_06_df['Ciclo'] == float(ciclo_nome.split('.')[0]) + float(ciclo_nome.split('.')[1])/10]
            if not ciclo_row.empty:
                # Procurar coluna de leads
                leads_col = None
                for col in excel_06_df.columns:
                    if 'leads' in col.lower() and 'matriz' not in col.lower():
                        leads_col = col
                        break
                if leads_col:
                    count_excel_06 = ciclo_row[leads_col].values[0]
        except:
            pass
    
    count_hubspot_hoje = hubspot_hoje.get(ciclo_nome)
    
    # Exibir comparação
    print(f"\n  Backup CSV (30/01):        {count_backup:>6,} leads")
    print(f"  Atual CSV (06/02):         {count_atual:>6,} leads")
    
    if count_excel_30 is not None:
        print(f"  Excel 30/01:               {int(count_excel_30):>6,} leads", end="")
        if int(count_excel_30) != count_backup:
            print(f" ⚠️ DIVERGÊNCIA: +{int(count_excel_30) - count_backup:+d}")
        else:
            print(" ✓")
    
    if count_excel_06 is not None:
        print(f"  Excel 06/02:               {int(count_excel_06):>6,} leads", end="")
        if int(count_excel_06) != count_atual:
            print(f" ⚠️ DIVERGÊNCIA: {int(count_excel_06) - count_atual:+d}")
        else:
            print(" ✓")
    
    if count_hubspot_hoje is not None:
        print(f"  HubSpot HOJE (filtro):     {count_hubspot_hoje:>6,} leads", end="")
        if count_hubspot_hoje != count_atual:
            diff_hoje = count_hubspot_hoje - count_atual
            print(f" ⚠️ DIVERGÊNCIA: {diff_hoje:+d}")
        else:
            print(" ✓")
    
    # Análise
    print(f"\n  Análise:")
    
    if count_excel_30 is not None and int(count_excel_30) != count_backup:
        print(f"    ❌ Excel 30/01 NÃO bate com Backup (diferença: {int(count_excel_30) - count_backup:+d})")
    
    if count_excel_06 is not None and int(count_excel_06) != count_atual:
        print(f"    ❌ Excel 06/02 NÃO bate com CSV Atual (diferença: {int(count_excel_06) - count_atual:+d})")
    
    if count_hubspot_hoje is not None and count_hubspot_hoje != count_atual:
        print(f"    ❌ HubSpot HOJE NÃO bate com CSV Atual (diferença: {count_hubspot_hoje - count_atual:+d})")
        print(f"       Possíveis razões:")
        print(f"       • Filtros diferentes no HubSpot")
        print(f"       • CSV tem registros deletados/incluídos")
        print(f"       • Timestamp de criação diferente")

print("\n\n" + "═" * 130)
print("LEITURA DOS ARQUIVOS EXCEL")
print("═" * 130 + "\n")

if excel_30_df is not None:
    print("Excel 30/01 - Colunas disponíveis:")
    print(excel_30_df.head())
    print(f"\nDados de cada ciclo:")
    print(excel_30_df.to_string())

print("\n" + "─" * 130 + "\n")

if excel_06_df is not None:
    print("Excel 06/02 - Colunas disponíveis:")
    print(excel_06_df.head())
    print(f"\nDados de cada ciclo:")
    print(excel_06_df.to_string())
