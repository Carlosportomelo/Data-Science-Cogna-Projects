import pandas as pd
from datetime import datetime
from pathlib import Path

print("=" * 160)
print("ANÁLISE COMPLETA DE CICLOS - LEADS E MATRÍCULAS")
print("=" * 160)
print()

# Localiza o arquivo de análise mais recente na pasta de outputs do projeto
project_root = Path(__file__).resolve().parents[2]
outputs_dir = project_root / 'outputs' / 'analise_consolidada_leads_matriculas'
analise_arquivo = None
if outputs_dir.exists():
    excels = sorted(outputs_dir.glob('analise_consolidada_leads_matriculas_*.xlsx'))
    if excels:
        analise_arquivo = excels[-1]

if analise_arquivo is None:
    # Fallback para caminho antigo (compatibilidade), mas informa o usuário
    analise_arquivo = Path(r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-01-30.xlsx')

print(f'Usando arquivo de análise: {analise_arquivo}')
df_analise_30 = pd.read_excel(analise_arquivo, sheet_name='Resumo e Insights')

print("DADOS DA ANÁLISE DO DIA 30/01/2026:")
print(df_analise_30.to_string(index=False))
print()
print()

# Carregar bases (usar paths relativos ao projeto)
base_backup = project_root / 'data' / 'backup' / 'hubspot_leads_backup.csv'
base_atual = project_root / 'data' / 'hubspot_leads.csv'

df_backup = pd.read_csv(base_backup)
df_atual = pd.read_csv(base_atual)

df_backup['Data de criação'] = pd.to_datetime(df_backup['Data de criação'], format='%Y-%m-%d %H:%M', errors='coerce')
df_backup['Data de fechamento'] = pd.to_datetime(df_backup['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')

df_atual['Data de criação'] = pd.to_datetime(df_atual['Data de criação'], format='%Y-%m-%d %H:%M', errors='coerce')
df_atual['Data de fechamento'] = pd.to_datetime(df_atual['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')

# Definir ciclos - baseado nos números que vemos na análise do dia 30
# Ciclos 23.1, 24.1, 25.1, 26.1

ciclos_defs = {}
for yy in ['23', '24', '25', '26']:
    label = f"{yy}.1"
    start_year = 2000 + int(yy) - 1
    start = pd.Timestamp(f"{start_year}-10-01")
    end = pd.Timestamp(f"{start_year + 1}-02-20") + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    ciclos_defs[label] = (start, end)

print("=" * 160)
print("COMPARAÇÃO POR CICLO - LEADS E MATRÍCULAS")
print("=" * 160)
print()

print(f"{'Ciclo':<8} | {'Ana30 Leads':<15} | {'Backup Leads':<15} | {'Dif':<10} | {'Ana30 Mat':<15} | {'Backup Mat':<15} | {'Dif':<10}")
print("-" * 160)

for ciclo, (data_inicio, data_fim) in ciclos_defs.items():
    # Filtrar leads
    df_ciclo_backup = df_backup[(df_backup['Data de criação'] >= data_inicio) & 
                                 (df_backup['Data de criação'] <= data_fim)]
    
    # Filtrar matrículas por data de fechamento
    df_mat_ciclo_backup = df_ciclo_backup[df_ciclo_backup['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)]
    df_mat_ciclo_backup = df_mat_ciclo_backup[(df_mat_ciclo_backup['Data de fechamento'] >= data_inicio) & 
                                               (df_mat_ciclo_backup['Data de fechamento'] <= data_fim)]
    
    qtd_leads = len(df_ciclo_backup)
    qtd_mat = len(df_mat_ciclo_backup)
    
    # Buscar na análise do dia 30
    ana30_row = df_analise_30[df_analise_30['Ciclo'].astype(str).str.startswith(ciclo.replace('.1', ''))]
    if len(ana30_row) > 0:
        ana30_leads = ana30_row.iloc[0]['Leads']
        ana30_mat = ana30_row.iloc[0]['Matriculas']
    else:
        ana30_leads = '?'
        ana30_mat = '?'
    
    dif_leads = f"{int(qtd_leads) - int(ana30_leads)}" if ana30_leads != '?' else 'N/A'
    dif_mat = f"{int(qtd_mat) - int(ana30_mat)}" if ana30_mat != '?' else 'N/A'
    
    print(f"{ciclo:<8} | {ana30_leads if ana30_leads != '?' else 'N/A':<15} | {qtd_leads:<15} | {dif_leads:<10} | {ana30_mat if ana30_mat != '?' else 'N/A':<15} | {qtd_mat:<15} | {dif_mat:<10}")

print()
print("=" * 160)
print("RESUMO DAS DISCREPÂNCIAS")
print("=" * 160)
print()

print("PROBLEMA ENCONTRADO:")
print("─" * 160)
print()
print("Ciclo 25.1: Matrículas diferem por ~200 (análise dia 30 tem 2.800, mas backup tem 2.595)")
print()
print("Hipótese:")
print("1. Análise do dia 30 pode estar contando matrículas com filtro diferente")
print("2. Ou estava usando base DIFERENTE na época")
print("3. Ou aplicando lógica especial para 're-certificar' matrículas")
print()
