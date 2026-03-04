import pandas as pd

excel_06 = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-02-06.xlsx'

print("=" * 150)
print("ANALISE DE TODAS AS SHEETS DO EXCEL 06/02".center(150))
print("=" * 150 + "\n")

# Carregar todas as sheets
xls = pd.ExcelFile(excel_06)

for sheet_name in xls.sheet_names:
    print(f"\n{'=' * 150}")
    print(f"SHEET: {sheet_name}")
    print(f"{'=' * 150}\n")
    
    df = pd.read_excel(excel_06, sheet_name=sheet_name)
    
    print(f"Dimensões: {df.shape[0]} linhas × {df.shape[1]} colunas")
    print(f"\nColunas: {list(df.columns)}\n")
    
    # Mostrar primeiras 10 linhas
    print(df.head(10).to_string())
    
    # Se tem "Ciclo" coluna, mostrar agrupado por ciclo
    if 'Ciclo' in df.columns:
        print(f"\n[*] Contagem por Ciclo:")
        if 'Leads' in df.columns:
            print(df.groupby('Ciclo')['Leads'].sum())
        elif 'quantidade' in df.columns or 'Quantidade' in df.columns:
            col = 'quantidade' if 'quantidade' in df.columns else 'Quantidade'
            print(df.groupby('Ciclo')[col].sum())

print("\n\n" + "=" * 150)
print("TOTAL DE LEADS POR CICLO (resumo)".center(150))
print("=" * 150 + "\n")

# Sheet Leads (Canais)
try:
    df_leads_canais = pd.read_excel(excel_06, sheet_name='Leads (Canais)')
    if 'Ciclo' in df_leads_canais.columns:
        print("Sheet 'Leads (Canais)':")
        ciclo_counts = df_leads_canais.groupby('Ciclo').size()
        for ciclo, count in ciclo_counts.items():
            print(f"  {ciclo}: {count:,}")
except:
    pass

print()

# Sheet Leads (Unidades)
try:
    df_leads_unidades = pd.read_excel(excel_06, sheet_name='Leads (Unidades)')
    if 'Ciclo' in df_leads_unidades.columns:
        print("Sheet 'Leads (Unidades)':")
        ciclo_counts = df_leads_unidades.groupby('Ciclo').size()
        for ciclo, count in ciclo_counts.items():
            print(f"  {ciclo}: {count:,}")
except:
    pass
import pandas as pd

excel_06 = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-02-06.xlsx'

print("=" * 150)
print("ANALISE DE TODAS AS SHEETS DO EXCEL 06/02".center(150))
print("=" * 150 + "\n")

# Carregar todas as sheets
xls = pd.ExcelFile(excel_06)

for sheet_name in xls.sheet_names:
    print(f"\n{'=' * 150}")
    print(f"SHEET: {sheet_name}")
    print(f"{'=' * 150}\n")
    
    df = pd.read_excel(excel_06, sheet_name=sheet_name)
    
    print(f"Dimensões: {df.shape[0]} linhas × {df.shape[1]} colunas")
    print(f"\nColunas: {list(df.columns)}\n")
    
    # Mostrar primeiras 10 linhas
    print(df.head(10).to_string())
    
    # Se tem "Ciclo" coluna, mostrar agrupado por ciclo
    if 'Ciclo' in df.columns:
        print(f"\n[*] Contagem por Ciclo:")
        if 'Leads' in df.columns:
            print(df.groupby('Ciclo')['Leads'].sum())
        elif 'quantidade' in df.columns or 'Quantidade' in df.columns:
            col = 'quantidade' if 'quantidade' in df.columns else 'Quantidade'
            print(df.groupby('Ciclo')[col].sum())

print("\n\n" + "=" * 150)
print("TOTAL DE LEADS POR CICLO (resumo)".center(150))
print("=" * 150 + "\n")

# Sheet Leads (Canais)
try:
    df_leads_canais = pd.read_excel(excel_06, sheet_name='Leads (Canais)')
    if 'Ciclo' in df_leads_canais.columns:
        print("Sheet 'Leads (Canais)':")
        ciclo_counts = df_leads_canais.groupby('Ciclo').size()
        for ciclo, count in ciclo_counts.items():
            print(f"  {ciclo}: {count:,}")
except:
    pass

print()

# Sheet Leads (Unidades)
try:
    df_leads_unidades = pd.read_excel(excel_06, sheet_name='Leads (Unidades)')
    if 'Ciclo' in df_leads_unidades.columns:
        print("Sheet 'Leads (Unidades)':")
        ciclo_counts = df_leads_unidades.groupby('Ciclo').size()
        for ciclo, count in ciclo_counts.items():
            print(f"  {ciclo}: {count:,}")
except:
    pass
