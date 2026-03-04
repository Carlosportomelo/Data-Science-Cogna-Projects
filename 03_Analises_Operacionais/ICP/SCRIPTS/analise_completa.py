import pandas as pd
import numpy as np
import warnings
import sys
warnings.filterwarnings('ignore')

# Configurar encoding UTF-8 para output
sys.stdout.reconfigure(encoding='utf-8')

# Carregar os dados
file_path = r'c:\Users\a483650\Projetos\03_Analises_Operacionais\ICP\DATA\Escolas_Novas_2025.xlsx'
df = pd.read_excel(file_path, sheet_name='Escolas')

# Limpar dados - remover linhas vazias
df_clean = df.dropna(subset=['Nome da Escola']).copy()

print("=" * 100)
print("ANÁLISE COMPLETA - ICP ESCOLAS 2025")
print("=" * 100)
print(f"\nTotal de escolas analisadas: {len(df_clean)}")
print("\n")

# Mostrar todos os dados disponíveis
print("=" * 100)
print("DADOS COMPLETOS DAS ESCOLAS")
print("=" * 100)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', 50)
print(df_clean[['Nome da Escola', 'Perfil', 'Alunado', 'Horas ', 'Seg', 'Faturamento', 'Origem', 'Inglês']].to_string(index=False))
print("\n")

# 1. TENDÊNCIA DE ESCOLAS COM TODOS OS ANOS CONTRATADOS
print("=" * 100)
print("1. ANÁLISE DE SEGMENTOS CONTRATADOS (Ed. Infantil, Anos Iniciais, Finais e E.M.)")
print("=" * 100)

# Verificar quais escolas têm informação sobre segmentos
df_clean['Seg_limpo'] = df_clean['Seg'].fillna('')
segmentos_total = []

for idx, row in df_clean.iterrows():
    seg = str(row['Seg']).lower()
    tem_infantil = 'infantil' in seg
    tem_fundamental = 'fundamental' in seg or 'fund' in seg
    tem_medio = 'médio' in seg or 'medio' in seg or 'ensino médio' in seg
    
    if tem_infantil and tem_fundamental and tem_medio:
        segmentos_total.append('Completo')
    else:
        segmentos_total.append('Parcial')

df_clean['Cobertura_Segmentos'] = segmentos_total

# Contar escolas com todos os segmentos
escolas_completas = df_clean[df_clean['Cobertura_Segmentos'] == 'Completo']
escolas_parciais = df_clean[df_clean['Cobertura_Segmentos'] == 'Parcial']

print(f"\nEscolas com TODOS os segmentos (Infantil + Fund + EM): {len(escolas_completas)} ({len(escolas_completas)/len(df_clean)*100:.1f}%)")
print(f"Escolas com segmentos PARCIAIS: {len(escolas_parciais)} ({len(escolas_parciais)/len(df_clean)*100:.1f}%)")

print("\n--- Escolas com cobertura COMPLETA ---")
for idx, row in escolas_completas.iterrows():
    print(f"  • {row['Nome da Escola']} - Segmentos: {row['Seg']}")

print("\n--- Escolas com cobertura PARCIAL ---")
for idx, row in escolas_parciais.iterrows():
    print(f"  • {row['Nome da Escola']} - Segmentos: {row['Seg']}")

# 2. COMPOSIÇÃO DA RECEITA (English Solutions vs Skies vs Compartilhado)
print("\n" + "=" * 100)
print("2. COMPOSIÇÃO DA RECEITA - ENGLISH SOLUTIONS vs SKIES vs COMPARTILHADO")
print("=" * 100)

# Analisar pela coluna de faturamento e outras informações
# Converter faturamento para numérico
df_clean['Faturamento_num'] = pd.to_numeric(df_clean['Faturamento'], errors='coerce')

# Tentar identificar pelo perfil ou origem
# Isso requer análise manual dos dados ou ter mais informações
print("\nNOTA: Para análise precisa da composição entre English Solutions e Skies,")
print("seria necessário ter uma coluna específica identificando o produto.")
print("\nAnálise baseada nos dados disponíveis:")
print(f"  • Faturamento Total: R$ {df_clean['Faturamento_num'].sum():,.2f}")
print(f"  • Ticket Médio por Escola: R$ {df_clean['Faturamento_num'].mean():,.2f}")
print(f"  • Mediana de Faturamento: R$ {df_clean['Faturamento_num'].median():,.2f}")

# 3. REGIÕES COM DESTAQUE
print("\n" + "=" * 100)
print("3. ANÁLISE POR ORIGEM DE CONTRATO")
print("=" * 100)

origem_analise = df_clean.groupby('Origem').agg({
    'Nome da Escola': 'count',
    'Faturamento_num': 'sum'
}).rename(columns={'Nome da Escola': 'Número de Escolas', 'Faturamento_num': 'Faturamento Total'})
origem_analise = origem_analise.sort_values('Faturamento Total', ascending=False)

print("\n--- Por Número de Escolas e Faturamento ---")
print(origem_analise.to_string())

print(f"\n\nPrincipal origem por NÚMERO DE ESCOLAS: {origem_analise['Número de Escolas'].idxmax()}")
print(f"Principal origem por FATURAMENTO: {origem_analise['Faturamento Total'].idxmax()}")

# 4. BROWNFIELD vs GREENFIELD
print("\n" + "=" * 100)
print( "4. ANÁLISE BROWNFIELD (com inglês anterior) vs GREENFIELD (sem inglês anterior)")
print("=" * 100)

# Classificar escolas
df_clean['Tipo_Campo'] = df_clean['Inglês'].apply(
    lambda x: 'Brownfield' if pd.notna(x) and 'não' not in str(x).lower() and 'sim' in str(x).lower() else 'Greenfield'
)

campo_analise = df_clean.groupby('Tipo_Campo').agg({
    'Nome da Escola': 'count',
    'Faturamento_num': 'sum',
    'Alunado': 'sum'
}).rename(columns={'Nome da Escola': 'Número de Escolas', 'Faturamento_num': 'Faturamento Total'})

print("\n--- Análise Brownfield vs Greenfield ---")
print(campo_analise.to_string())

print("\n--- Escolas por Tipo ---")
for tipo in ['Brownfield', 'Greenfield']:
    print(f"\n{tipo.upper()}:")
    escolas_tipo = df_clean[df_clean['Tipo_Campo'] == tipo]
    for idx, row in escolas_tipo.iterrows():
        print(f"  • {row['Nome da Escola']} - Inglês anterior: {row['Inglês']}")

# 5. RELAÇÃO ENTRE INGLÊS ANTERIOR E HORAS CONTRATADAS
print("\n" + "=" * 100)
print("5. RELAÇÃO: INGLÊS ANTERIOR vs HORAS CONTRATADAS (>3)")
print("=" * 100)

# Extrair número de horas
def extrair_horas(texto):
    if pd.isna(texto):
        return None
    texto_str = str(texto).lower()
    if '5' in texto_str:
        return 5
    elif '4' in texto_str:
        return 4
    elif '3' in texto_str:
        return 3
    elif '2' in texto_str:
        return 2
    return None

df_clean['Horas_num'] = df_clean['Horas '].apply(extrair_horas)
df_clean['Mais_3_horas'] = df_clean['Horas_num'].apply(lambda x: 'Sim (>3h)' if pd.notna(x) and x > 3 else 'Não (≤3h)' if pd.notna(x) else 'N/D')

# Análise cruzada
cross_analise = pd.crosstab(
    df_clean['Tipo_Campo'], 
    df_clean['Mais_3_horas'], 
    margins=True
)

print("\n--- Cruzamento: Tipo de Campo vs Horas Contratadas ---")
print(cross_analise.to_string())

# Estatísticas mais detalhadas
print("\n--- Detalhamento ---")
for tipo in ['Brownfield', 'Greenfield']:
    escolas_tipo = df_clean[df_clean['Tipo_Campo'] == tipo]
    escolas_mais_3 = escolas_tipo[escolas_tipo['Horas_num'] > 3]
    if len(escolas_tipo[escolas_tipo['Horas_num'].notna()]) > 0:
        perc = len(escolas_mais_3) / len(escolas_tipo[escolas_tipo['Horas_num'].notna()]) * 100
        print(f"{tipo}: {len(escolas_mais_3)}/{len(escolas_tipo[escolas_tipo['Horas_num'].notna()])} escolas com >3h ({perc:.1f}%)")

# 6. PERFIL COM MAIS RENTABILIDADE E LTV
print("\n" + "=" * 100)
print("6. ANÁLISE DE PERFIL: RENTABILIDADE E LTV")
print("=" * 100)

perfil_analise = df_clean.groupby('Perfil').agg({
    'Nome da Escola': 'count',
    'Faturamento_num': ['sum', 'mean', 'median'],
    'Alunado': ['sum', 'mean']
}).round(2)

perfil_analise.columns = ['Qtd Escolas', 'Fat. Total', 'Fat. Médio', 'Fat. Mediano', 'Alunos Total', 'Alunos Médio']

print("\n--- Análise por Perfil de Escola ---")
print(perfil_analise.to_string())

print("\n--- Ranking por Faturamento Médio (proxy para LTV/Rentabilidade) ---")
ranking = perfil_analise.sort_values('Fat. Médio', ascending=False)
for idx, (perfil, dados) in enumerate(ranking.iterrows(), 1):
    print(f"{idx}. {perfil}: R$ {dados['Fat. Médio']:,.2f} (média) | {int(dados['Qtd Escolas'])} escolas")

# 7. PERFIL DAS ESCOLAS QUE CAUSAM CHURN
print("\n" + "=" * 100)
print("7. ANÁLISE DE CHURN - PERFIL E PORTE")
print("=" * 100)
print("\nNOTA: Os dados fornecidos são de ESCOLAS NOVAS (ganhos) em 2025.")
print("Para análise de CHURN completa, seria necessário ter dados de:")
print("  • Escolas que cancelaram/saíram")
print("  • Motivos de saída")
print("  • Histórico de relacionamento")
print("\nCom os dados atuais, podemos inferir PERFIS DE RISCO analisando características:")

# Classificar porte por alunado
def classificar_porte(alunos):
    if pd.isna(alunos):
        return 'N/D'
    if alunos < 300:
        return 'Pequeno (< 300)'
    elif alunos < 600:
        return 'Médio (300-600)'
    else:
        return 'Grande (> 600)'

df_clean['Porte'] = df_clean['Alunado'].apply(classificar_porte)

porte_perfil = df_clean.groupby(['Porte', 'Perfil']).agg({
    'Nome da Escola': 'count',
    'Faturamento_num': 'mean'
}).rename(columns={'Nome da Escola': 'Qtd', 'Faturamento_num': 'Fat. Médio'})

print("\n--- Distribuição por Porte e Perfil ---")
print(porte_perfil.to_string())

# 8. ESCOLAS PEDAGÓGICAS PARA UPSELL
print("\n" + "=" * 100)
print("8. ESCOLAS PEDAGÓGICAS COM POTENCIAL DE UPSELL - LISTA DE PRIORIDADES")
print("=" * 100)

# Filtrar escolas pedagógicas
escolas_pedagogicas = df_clean[df_clean['Perfil'] == 'Pedagógica'].copy()

# Critérios de priorização:
# 1. Segmentos parciais (pode expandir para mais anos)
# 2. Menos de 4 horas (pode aumentar carga horária)
# 3. Alto alunado (mais potencial de receita)
# 4. Já tem relacionamento positivo

escolas_pedagogicas['Score_Upsell'] = 0

# Pontuação por segmentos parciais
escolas_pedagogicas.loc[escolas_pedagogicas['Cobertura_Segmentos'] == 'Parcial', 'Score_Upsell'] += 3

# Pontuação por horas baixas
escolas_pedagogicas.loc[escolas_pedagogicas['Horas_num'] <= 3, 'Score_Upsell'] += 2

# Pontuação por alunado alto
escolas_pedagogicas.loc[escolas_pedagogicas['Alunado'] > 500, 'Score_Upsell'] += 3
escolas_pedagogicas.loc[(escolas_pedagogicas['Alunado'] > 300) & (escolas_pedagogicas['Alunado'] <= 500), 'Score_Upsell'] += 2

# Ordenar por score
escolas_pedagogicas_sorted = escolas_pedagogicas.sort_values('Score_Upsell', ascending=False)

print("\n--- LISTA PRIORIZADA DE UPSELL (Escolas Pedagógicas) ---\n")
for idx, (_, row) in enumerate(escolas_pedagogicas_sorted.iterrows(), 1):
    print(f"{idx}. {row['Nome da Escola']} [Score: {row['Score_Upsell']}]")
    print(f"   • Alunado: {int(row['Alunado']) if pd.notna(row['Alunado']) else 'N/D'} alunos")
    print(f"   • Faturamento Atual: R$ {row['Faturamento_num']:,.2f}")
    print(f"   • Horas: {row['Horas ']} ({row['Horas_num']}h)")
    print(f"   • Segmentos: {row['Seg']} [{row['Cobertura_Segmentos']}]")
    print(f"   • Oportunidades:")
    
    oportunidades = []
    if row['Cobertura_Segmentos'] == 'Parcial':
        oportunidades.append("Expandir para mais segmentos educacionais")
    if pd.notna(row['Horas_num']) and row['Horas_num'] <= 3:
        oportunidades.append("Aumentar carga horária (de {} para 4-5h)".format(row['Horas_num']))
    if row['Alunado'] > 500:
        oportunidades.append("Alto volume de alunos = alto potencial de receita adicional")
    
    for op in oportunidades:
        print(f"     → {op}")
    print()

# RESUMO EXECUTIVO
print("\n" + "=" * 100)
print("RESUMO EXECUTIVO - PRINCIPAIS INSIGHTS")
print("=" * 100)
print(f"""
📊 VISÃO GERAL:
   • Total de Escolas: {len(df_clean)}
   • Faturamento Total: R$ {df_clean['Faturamento_num'].sum():,.2f}
   • Ticket Médio: R$ {df_clean['Faturamento_num'].mean():,.2f}
   • Alunos Totais: {int(df_clean['Alunado'].sum())}

🎯 SEGMENTOS:
   • Escolas com cobertura completa (Infantil + Fund + EM): {len(escolas_completas)} ({len(escolas_completas)/len(df_clean)*100:.1f}%)
   • Escolas com cobertura parcial: {len(escolas_parciais)} ({len(escolas_parciais)/len(df_clean)*100:.1f}%)

🏗️ BROWNFIELD vs GREENFIELD:
   • Brownfield (com inglês anterior): {len(df_clean[df_clean['Tipo_Campo'] == 'Brownfield'])} escolas
   • Greenfield (sem inglês anterior): {len(df_clean[df_clean['Tipo_Campo'] == 'Greenfield'])} escolas

💰 PERFIL MAIS RENTÁVEL:
   • {ranking.index[0]}: R$ {ranking.iloc[0]['Fat. Médio']:,.2f} (faturamento médio)

🎓 TOP OPORTUNIDADES DE UPSELL:
""")

for idx, (_, row) in enumerate(escolas_pedagogicas_sorted.head(3).iterrows(), 1):
    print(f"   {idx}. {row['Nome da Escola']} - Score: {row['Score_Upsell']}")

print("\n" + "=" * 100)
