import pandas as pd

print("=" * 160)
print("TESTE: CICLOS PASSSADOS COM CORTE EM 30/01/2026")
print("=" * 160)
print()

# Carregar análise do dia 30
analise_arquivo = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-01-30.xlsx'
df_analise_30 = pd.read_excel(analise_arquivo, sheet_name='Resumo e Insights')

print("DADOS DA ANÁLISE DO DIA 30:")
print(df_analise_30[['Ciclo', 'Leads', 'Matriculas']].to_string(index=False))
print()

# Carregar base backup
base_backup = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads_backup.csv'
df_backup = pd.read_csv(base_backup)

df_backup['Data de criação'] = pd.to_datetime(df_backup['Data de criação'], format='%Y-%m-%d %H:%M', errors='coerce')
df_backup['Data de fechamento'] = pd.to_datetime(df_backup['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')

# Ciclos com final fixo em 30/01/2026
ciclos_com_corte_30jan = {
    '23.1': (pd.Timestamp('2022-10-01'), pd.Timestamp('2023-01-30 23:59:59')),  # Oct 22 - Jan 23
    '24.1': (pd.Timestamp('2023-10-01'), pd.Timestamp('2024-01-30 23:59:59')),  # Oct 23 - Jan 24
    '25.1': (pd.Timestamp('2024-10-01'), pd.Timestamp('2025-01-30 23:59:59')),  # Oct 24 - Jan 25
    '26.1': (pd.Timestamp('2025-10-01'), pd.Timestamp('2026-01-30 23:59:59')),  # Oct 25 - Jan 26
}

print("=" * 160)
print("TESTE 1: CICLOS COM CORTE EM 30/01 (4 MESES cada)")
print("=" * 160)
print()

print(f"{'Ciclo':<8} | {'Ana30 Leads':<15} | {'Backup Leads':<15} | {'Dif':<10} | {'Ana30 Mat':<15} | {'Backup Mat':<15} | {'Dif':<10} | {'Match?':<10}")
print("-" * 160)

for ciclo, (data_inicio, data_fim) in ciclos_com_corte_30jan.items():
    # Filtrar leads
    df_ciclo_backup = df_backup[(df_backup['Data de criação'] >= data_inicio) & 
                                 (df_backup['Data de criação'] <= data_fim)]
    
    # Filtrar matrículas
    df_mat_backup = df_ciclo_backup[df_ciclo_backup['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)]
    df_mat_backup = df_mat_backup[(df_mat_backup['Data de fechamento'] >= data_inicio) & 
                                   (df_mat_backup['Data de fechamento'] <= data_fim)]
    
    qtd_leads_backup = len(df_ciclo_backup)
    qtd_mat_backup = len(df_mat_backup)
    
    # Comparar com análise
    ana30_row = df_analise_30[df_analise_30['Ciclo'] == float(ciclo)]
    if len(ana30_row) > 0:
        ana30_leads = ana30_row.iloc[0]['Leads']
        ana30_mat = ana30_row.iloc[0]['Matriculas']
    else:
        ana30_leads = -1
        ana30_mat = -1
    
    dif_leads = qtd_leads_backup - ana30_leads if ana30_leads > 0 else -999
    dif_mat = qtd_mat_backup - ana30_mat if ana30_mat > 0 else -999
    
    match_leads = "✓" if abs(dif_leads) < 5 else "✗"
    
    print(f"{ciclo:<8} | {ana30_leads:<15} | {qtd_leads_backup:<15} | {dif_leads:<10} | {ana30_mat:<15} | {qtd_mat_backup:<15} | {dif_mat:<10} | {match_leads:<10}")

print()
print("=" * 160)
print("TESTE 2: CICLOS OUT A JAN (5 MESES cada)")
print("=" * 160)
print()

ciclos_out_jan = {
    '23.1': (pd.Timestamp('2022-10-01'), pd.Timestamp('2023-01-31 23:59:59')),
    '24.1': (pd.Timestamp('2023-10-01'), pd.Timestamp('2024-01-31 23:59:59')),
    '25.1': (pd.Timestamp('2024-10-01'), pd.Timestamp('2025-01-31 23:59:59')),
    '26.1': (pd.Timestamp('2025-10-01'), pd.Timestamp('2026-01-31 23:59:59')),
}

print(f"{'Ciclo':<8} | {'Ana30 Leads':<15} | {'Backup Leads':<15} | {'Dif':<10} | {'Match?':<10}")
print("-" * 160)

for ciclo, (data_inicio, data_fim) in ciclos_out_jan.items():
    df_ciclo_backup = df_backup[(df_backup['Data de criação'] >= data_inicio) & 
                                 (df_backup['Data de criação'] <= data_fim)]
    
    df_mat_backup = df_ciclo_backup[df_ciclo_backup['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)]
    df_mat_backup = df_mat_backup[(df_mat_backup['Data de fechamento'] >= data_inicio) & 
                                   (df_mat_backup['Data de fechamento'] <= data_fim)]
    
    qtd_leads = len(df_ciclo_backup)
    qtd_mat = len(df_mat_backup)
    
    ana30_row = df_analise_30[df_analise_30['Ciclo'] == float(ciclo)]
    if len(ana30_row) > 0:
        ana30_leads = ana30_row.iloc[0]['Leads']
        ana30_mat = ana30_row.iloc[0]['Matriculas']
    else:
        ana30_leads = -1
        ana30_mat = -1
    
    dif_leads = qtd_leads - ana30_leads if ana30_leads > 0 else -999
    match = "✓" if abs(dif_leads) < 5 else "✗"
    
    print(f"{ciclo:<8} | {ana30_leads:<15} | {qtd_leads:<15} | {dif_leads:<10} | {match:<10}")

print()
print("=" * 160)
print("RESUMO")
print("=" * 160)
print()
print("Se os números baterem com 'corte em 30/01', então a análise do dia 30")
print("foi feita com data de corte específica, não cobriu os ciclos completos.")
