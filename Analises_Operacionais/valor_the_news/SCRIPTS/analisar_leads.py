import pandas as pd
from datetime import datetime

# Carregar dados de 2026
df_2026 = pd.read_csv(r"c:\Users\a483650\Projetos\Analise_performance_valor_the_news\DATA\hubspot_leads_26.csv")

# Carregar dados de 2025
df_2025 = pd.read_csv(r"c:\Users\a483650\Projetos\Analise_performance_valor_the_news\DATA\hubspot_leads_25.csv")

# Converter datas
df_2026['Data de criação'] = pd.to_datetime(df_2026['Data de criação'])
df_2025['Data de criação'] = pd.to_datetime(df_2025['Data de criação'])

# Definir períodos
# Semana da ação 2026: 19-25 janeiro 2026
# Semana equivalente 2025: 19-25 janeiro 2025

inicio_2026 = '2026-01-19'
fim_2026 = '2026-01-25 23:59:59'

inicio_2025 = '2025-01-19'
fim_2025 = '2025-01-25 23:59:59'

# Filtrar semanas
leads_semana_2026 = df_2026[(df_2026['Data de criação'] >= inicio_2026) & (df_2026['Data de criação'] <= fim_2026)]
leads_semana_2025 = df_2025[(df_2025['Data de criação'] >= inicio_2025) & (df_2025['Data de criação'] <= fim_2025)]

print("="*60)
print("COMPARATIVO DE LEADS: SEMANA DA AÇÃO vs 2025")
print("="*60)

print(f"\nPeriodo 2026: 19/01 a 25/01/2026 (Semana da Acao)")
print(f"Periodo 2025: 19/01 a 25/01/2025 (Semana Equivalente)")

print(f"\n{'='*60}")
print("VOLUME TOTAL DE LEADS")
print("="*60)
total_2026 = len(leads_semana_2026)
total_2025 = len(leads_semana_2025)
crescimento = ((total_2026 - total_2025) / total_2025 * 100) if total_2025 > 0 else 0

print(f"Leads 2025: {total_2025}")
print(f"Leads 2026: {total_2026}")
print(f"Crescimento: +{crescimento:.2f}%")
print(f"Leads extras: +{total_2026 - total_2025}")

# Análise por dia
print(f"\n{'='*60}")
print("LEADS POR DIA")
print("="*60)

leads_por_dia_2026 = leads_semana_2026.groupby(leads_semana_2026['Data de criação'].dt.date).size()
leads_por_dia_2025 = leads_semana_2025.groupby(leads_semana_2025['Data de criação'].dt.date).size()

print("\n2026 (Semana da Ação):")
for data, qtd in leads_por_dia_2026.items():
    print(f"  {data}: {qtd} leads")

print("\n2025 (Semana Equivalente):")
for data, qtd in leads_por_dia_2025.items():
    print(f"  {data}: {qtd} leads")

# Análise por fonte de tráfego
print(f"\n{'='*60}")
print("LEADS POR FONTE DE TRÁFEGO")
print("="*60)

fonte_2026 = leads_semana_2026['Fonte original do tráfego'].value_counts()
fonte_2025 = leads_semana_2025['Fonte original do tráfego'].value_counts()

print("\n2026:")
for fonte, qtd in fonte_2026.items():
    print(f"  {fonte}: {qtd}")

print("\n2025:")
for fonte, qtd in fonte_2025.items():
    print(f"  {fonte}: {qtd}")

# Análise por etapa do negócio
print(f"\n{'='*60}")
print("LEADS POR ETAPA DO NEGÓCIO (2026)")
print("="*60)

etapa_2026 = leads_semana_2026['Etapa do negócio'].value_counts()
for etapa, qtd in etapa_2026.items():
    print(f"  {etapa}: {qtd}")

# Comparativo consolidado para o relatório
print(f"\n{'='*60}")
print("RESUMO PARA RELATÓRIO")
print("="*60)

# Criar DataFrame comparativo
comparativo = pd.DataFrame({
    'Métrica': ['Total de Leads', 'Leads/Dia (média)', 'Social Pago', 'Pesquisa Paga', 'Tráfego Direto'],
    '2025 (19-25/01)': [
        total_2025,
        round(total_2025/7, 1),
        fonte_2025.get('Social pago', 0),
        fonte_2025.get('Pesquisa paga', 0),
        fonte_2025.get('Tráfego direto', 0)
    ],
    '2026 (19-25/01)': [
        total_2026,
        round(total_2026/7, 1),
        fonte_2026.get('Social pago', 0),
        fonte_2026.get('Pesquisa paga', 0),
        fonte_2026.get('Tráfego direto', 0)
    ]
})

# Calcular variação
comparativo['Variação %'] = ((comparativo['2026 (19-25/01)'] - comparativo['2025 (19-25/01)']) / comparativo['2025 (19-25/01)'] * 100).round(2)
comparativo['Variação %'] = comparativo['Variação %'].apply(lambda x: f"+{x}%" if x > 0 else f"{x}%")

print(comparativo.to_string(index=False))

# Salvar dados para uso no relatório
comparativo.to_csv(r"c:\Users\a483650\Projetos\Analise_performance_valor_the_news\OUTPUTS\comparativo_leads.csv", index=False)
print(f"\nDados salvos em: OUTPUTS/comparativo_leads.csv")
