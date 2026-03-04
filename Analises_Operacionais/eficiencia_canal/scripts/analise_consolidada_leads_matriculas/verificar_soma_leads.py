import pandas as pd

excel_06 = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-02-06.xlsx'

# Carregar dados brutos
df_leads_canais = pd.read_excel(excel_06, sheet_name='Leads (Canais)')
df_resumo = pd.read_excel(excel_06, sheet_name='Resumo e Insights')

print("=" * 120)
print("VERIFICACAO: Totais em 'Leads (Canais)' vs 'Resumo e Insights'")
print("=" * 120 + "\n")

print("Sheet 'Leads (Canais)' - primeiras linhas:")
print(df_leads_canais.head(12))

print("\n\nSomando por ciclo em 'Leads (Canais)':\n")

for ciclo in ['23.1', '24.1', '25.1', '26.1']:
    total_por_ciclo = df_leads_canais[ciclo].sum()
    resumo_value = df_resumo[df_resumo['Ciclo'] == float(ciclo)]['Leads'].values[0]
    
    print(f"Ciclo {ciclo}:")
    print(f"  Soma em 'Leads (Canais)': {total_por_ciclo:,.0f}")
    print(f"  Valor em 'Resumo':        {resumo_value:,.0f}")
    
    if total_por_ciclo != resumo_value:
        print(f"  DISCREPANCIA: {resumo_value - total_por_ciclo:+.0f}")
    print()

print("\n" + "=" * 120)
print("CALCULO INVESTIGATIVO - Qual é a fonte do 10,527 para ciclo 25.1?")
print("=" * 120 + "\n")

# Ver detalhes do ciclo 25.1
print("Breakdown de Leads 25.1 por Canal:")
print(df_leads_canais[['Fonte original do tráfego', '25.1']])

print(f"\nTotal: {df_leads_canais['25.1'].sum():,.0f}")
print(f"Valor no Resumo: 10,527")
print(f"Diferenca: {10527 - df_leads_canais['25.1'].sum():+,.0f}")

# Se a diferenca for negativa, significa que o Resumo tem MAIS do que os canais somam
# Isso seria manipulacao!
