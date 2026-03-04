import pandas as pd

# Carregar análise do dia 30
analise_arquivo = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-01-30.xlsx'
df_analise_30 = pd.read_excel(analise_arquivo, sheet_name='Resumo e Insights')

print("Dados da análise 30:")
print(df_analise_30)
print()
print("Colunas:", df_analise_30.columns.tolist())
print("Ciclos:", df_analise_30['Ciclo'].tolist())
print("Tipos:", df_analise_30.dtypes)
print()

# Mostrar valor bruto
print("Primeiro ciclo:", repr(df_analise_30['Ciclo'].iloc[0]))
print("Tipo:", type(df_analise_30['Ciclo'].iloc[0]))
