import pandas as pd

print("=" * 160)
print("ANÁLISE COM CICLOS CORRETOS - out a mar (6 meses)")
print("=" * 160)
print()

# Carregar análise do dia 30
analise_arquivo = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-01-30.xlsx'
df_analise_30 = pd.read_excel(analise_arquivo, sheet_name='Resumo e Insights')

print("DADOS DA ANÁLISE DO DIA 30:")
print(df_analise_30[['Ciclo', 'Leads', 'Matriculas']].to_string(index=False))
print()

# Carregar bases
base_backup = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads_backup.csv'
base_atual = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\hubspot_leads.csv'

df_backup = pd.read_csv(base_backup)
df_atual = pd.read_csv(base_atual)

df_backup['Data de criação'] = pd.to_datetime(df_backup['Data de criação'], format='%Y-%m-%d %H:%M', errors='coerce')
df_backup['Data de fechamento'] = pd.to_datetime(df_backup['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')

df_atual['Data de criação'] = pd.to_datetime(df_atual['Data de criação'], format='%Y-%m-%d %H:%M', errors='coerce')
df_atual['Data de fechamento'] = pd.to_datetime(df_atual['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')

# Ciclos CORRETOS (out a mar, 6 meses)
ciclos_corretos = {
    '23.1': (pd.Timestamp('2022-10-01'), pd.Timestamp('2023-03-31 23:59:59')),
    '24.1': (pd.Timestamp('2023-10-01'), pd.Timestamp('2024-03-31 23:59:59')),
    '25.1': (pd.Timestamp('2024-10-01'), pd.Timestamp('2025-03-31 23:59:59')),
    '26.1': (pd.Timestamp('2025-10-01'), pd.Timestamp('2026-03-31 23:59:59')),
}

print("=" * 160)
print("COMPARAÇÃO CORRIGIDA - CICLOS OUT A MAR (6 MESES)")
print("=" * 160)
print()

print(f"{'Ciclo':<8} | {'Ana30 Leads':<15} | {'Backup Leads':<15} | {'Dif':<10} | {'Ana30 Mat':<15} | {'Backup Mat':<15} | {'Dif':<10} | {'Status':<20}")
print("-" * 160)

for ciclo, (data_inicio, data_fim) in ciclos_corretos.items():
    # Filtrar leads por data de criação
    df_ciclo_backup = df_backup[(df_backup['Data de criação'] >= data_inicio) & 
                                 (df_backup['Data de criação'] <= data_fim)]
    df_ciclo_atual = df_atual[(df_atual['Data de criação'] >= data_inicio) & 
                               (df_atual['Data de criação'] <= data_fim)]
    
    # Filtrar matrículas por data de fechamento
    df_mat_backup = df_ciclo_backup[df_ciclo_backup['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)]
    df_mat_backup = df_mat_backup[(df_mat_backup['Data de fechamento'] >= data_inicio) & 
                                   (df_mat_backup['Data de fechamento'] <= data_fim)]
    
    df_mat_atual = df_ciclo_atual[df_ciclo_atual['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)]
    df_mat_atual = df_mat_atual[(df_mat_atual['Data de fechamento'] >= data_inicio) & 
                                 (df_mat_atual['Data de fechamento'] <= data_fim)]
    
    qtd_leads_backup = len(df_ciclo_backup)
    qtd_leads_atual = len(df_ciclo_atual)
    qtd_mat_backup = len(df_mat_backup)
    qtd_mat_atual = len(df_mat_atual)
    
    # Buscar na análise do dia 30
    try:
        ciclo_float = float(ciclo)
        ana30_row = df_analise_30[df_analise_30['Ciclo'] == ciclo_float]
        if len(ana30_row) > 0:
            ana30_leads = ana30_row.iloc[0]['Leads']
            ana30_mat = ana30_row.iloc[0]['Matriculas']
        else:
            ana30_leads = -1
            ana30_mat = -1
    except:
        ana30_leads = -1
        ana30_mat = -1
    
    dif_leads_backup = qtd_leads_backup - ana30_leads if ana30_leads > 0 else 0
    dif_mat_backup = qtd_mat_backup - ana30_mat if ana30_mat > 0 else 0
    
    status_leads = "✓ OK" if abs(dif_leads_backup) < 5 else "✗ DIFERENÇA"
    
    print(f"{ciclo:<8} | {ana30_leads:<15} | {qtd_leads_backup:<15} | {dif_leads_backup:<10} | {ana30_mat:<15} | {qtd_mat_backup:<15} | {dif_mat_backup:<10} | {status_leads:<20}")

print()
print("=" * 160)
print("CONCLUSÃO")
print("=" * 160)
print()
print("Com os ciclos CORRETOS (out a mar), os números devem bater com a análise do dia 30.")
