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
print("ANÁLISE COMPLETA - ICP ESCOLAS 2025 (VERSÃO CORRIGIDA)")
print("=" * 100)
print(f"\nTotal de escolas analisadas: {len(df_clean)}")
print("\n")

# Processar dados
# 1. Classificar segmentos
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

# 2. Classificar Brownfield vs Greenfield - CLASSIFICAÇÃO CORRETA
# Brownfield = cidades COM franqueado Red Balloon (origem menciona franqueado/franquia)
# Greenfield = cidades SEM franqueado Red Balloon
df_clean['Tipo_Campo'] = df_clean['Origem'].apply(
    lambda x: 'Brownfield' if pd.notna(x) and ('franqueado' in str(x).lower() or 'franquia' in str(x).lower()) else 'Greenfield'
)

# 3. Classificar se tinha inglês anterior
df_clean['Tinha_Ingles_Anterior'] = df_clean['Inglês'].apply(
    lambda x: 'Sim' if pd.notna(x) and 'sim' in str(x).lower() else 'Não/Insatisfeito' if pd.notna(x) else 'N/D'
)

# 4. Extrair horas
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

df_clean['Horas_Numerico'] = df_clean['Horas '].apply(extrair_horas)
df_clean['Mais_3_horas'] = df_clean['Horas_Numerico'].apply(lambda x: 'Sim' if pd.notna(x) and x > 3 else 'Não' if pd.notna(x) else 'N/D')

# 5. Converter faturamento para numérico
df_clean['Faturamento_Numerico'] = pd.to_numeric(df_clean['Faturamento'], errors='coerce')

# 6. Classificar porte
def classificar_porte(alunos):
    if pd.isna(alunos):
        return 'N/D'
    if alunos < 300:
        return 'Pequeno'
    elif alunos < 600:
        return 'Médio'
    else:
        return 'Grande'

df_clean['Porte'] = df_clean['Alunado'].apply(classificar_porte)

# ANÁLISE 4: BROWNFIELD vs GREENFIELD (CORRIGIDO)
print("=" * 100)
print("4. ANÁLISE BROWNFIELD vs GREENFIELD (CORRIGIDA)")
print("=" * 100)
print("\nDEFINIÇÕES:")
print("  • Brownfield = Cidades COM franqueado Red Balloon (franqueado dá suporte)")
print("  • Greenfield = Cidades SEM franqueado Red Balloon\n")

campo_analise = df_clean.groupby('Tipo_Campo').agg({
    'Nome da Escola': 'count',
    'Faturamento_Numerico': 'sum',
    'Alunado': 'sum'
}).rename(columns={'Nome da Escola': 'Número de Escolas', 'Faturamento_Numerico': 'Faturamento Total'})

print("--- Análise Brownfield vs Greenfield ---")
print(campo_analise.to_string())
print("\n")

brownfield_count = len(df_clean[df_clean['Tipo_Campo'] == 'Brownfield'])
greenfield_count = len(df_clean[df_clean['Tipo_Campo'] == 'Greenfield'])
brownfield_perc = (brownfield_count / len(df_clean)) * 100
greenfield_perc = (greenfield_count / len(df_clean)) * 100

brownfield_fat = df_clean[df_clean['Tipo_Campo'] == 'Brownfield']['Faturamento_Numerico'].sum()
greenfield_fat = df_clean[df_clean['Tipo_Campo'] == 'Greenfield']['Faturamento_Numerico'].sum()

print(f"BROWNFIELD (com franqueado): {brownfield_count} escolas ({brownfield_perc:.1f}%) - R$ {brownfield_fat:,.2f}")
print(f"GREENFIELD (sem franqueado): {greenfield_count} escolas ({greenfield_perc:.1f}%) - R$ {greenfield_fat:,.2f}")

print("\n--- Escolas por Tipo ---")
for tipo in ['Brownfield', 'Greenfield']:
    print(f"\n{tipo.upper()}:")
    escolas_tipo = df_clean[df_clean['Tipo_Campo'] == tipo]
    for idx, row in escolas_tipo.iterrows():
        print(f"  • {row['Nome da Escola']} - Origem: {row['Origem']}")

# ANÁLISE ADICIONAL: INGLÊS ANTERIOR
print("\n" + "=" * 100)
print("ANÁLISE COMPLEMENTAR: INGLÊS ANTERIOR")
print("=" * 100)

ingles_analise = df_clean.groupby('Tinha_Ingles_Anterior').agg({
    'Nome da Escola': 'count',
    'Faturamento_Numerico': 'sum',
    'Alunado': 'sum'
}).rename(columns={'Nome da Escola': 'Número de Escolas', 'Faturamento_Numerico': 'Faturamento Total'})

print("\n--- Escolas por Inglês Anterior ---")
print(ingles_analise.to_string())
print("\n")

tinha_ingles = len(df_clean[df_clean['Tinha_Ingles_Anterior'] == 'Sim'])
nao_tinha = len(df_clean[df_clean['Tinha_Ingles_Anterior'] != 'Sim'])

print(f"Escolas que TINHAM inglês anterior: {tinha_ingles} ({tinha_ingles/len(df_clean)*100:.1f}%)")
print(f"Escolas que NÃO tinham ou estavam insatisfeitas: {nao_tinha} ({nao_tinha/len(df_clean)*100:.1f}%)")

# ANÁLISE 5: RELAÇÃO BROWNFIELD/GREENFIELD vs HORAS (CORRIGIDA)
print("\n" + "=" * 100)
print("5. RELAÇÃO: BROWNFIELD/GREENFIELD vs HORAS CONTRATADAS (>3h)")
print("=" * 100)

cross_analise = pd.crosstab(
    df_clean['Tipo_Campo'], 
    df_clean['Mais_3_horas'], 
    margins=True
)

print("\n--- Cruzamento: Brownfield/Greenfield vs Horas Contratadas ---")
print(cross_analise.to_string())
print("\n")

# Estatísticas detalhadas
for tipo in ['Brownfield', 'Greenfield']:
    escolas_tipo = df_clean[df_clean['Tipo_Campo'] == tipo]
    escolas_mais_3 = escolas_tipo[escolas_tipo['Horas_Numerico'] > 3]
    escolas_validas = escolas_tipo[escolas_tipo['Horas_Numerico'].notna()]
    if len(escolas_validas) > 0:
        perc = len(escolas_mais_3) / len(escolas_validas) * 100
        print(f"{tipo}: {len(escolas_mais_3)}/{len(escolas_validas)} escolas com >3h ({perc:.1f}%)")

# ANÁLISE 5B: RELAÇÃO INGLÊS ANTERIOR vs HORAS
print("\n" + "=" * 100)
print("5B. RELAÇÃO: INGLÊS ANTERIOR vs HORAS CONTRATADAS (>3h)")
print("=" * 100)

cross_ingles = pd.crosstab(
    df_clean['Tinha_Ingles_Anterior'], 
    df_clean['Mais_3_horas'], 
    margins=True
)

print("\n--- Cruzamento: Inglês Anterior vs Horas Contratadas ---")
print(cross_ingles.to_string())
print("\n")

# Estatísticas detalhadas
for tipo_ing in ['Sim', 'Não/Insatisfeito']:
    escolas_tipo = df_clean[df_clean['Tinha_Ingles_Anterior'] == tipo_ing]
    escolas_mais_3 = escolas_tipo[escolas_tipo['Horas_Numerico'] > 3]
    escolas_validas = escolas_tipo[escolas_tipo['Horas_Numerico'].notna()]
    if len(escolas_validas) > 0:
        perc = len(escolas_mais_3) / len(escolas_validas) * 100
        print(f"{tipo_ing}: {len(escolas_mais_3)}/{len(escolas_validas)} escolas com >3h ({perc:.1f}%)")

# ANÁLISE 6: PERFIL MAIS RENTÁVEL
print("\n" + "=" * 100)
print("6. ANÁLISE DE PERFIL: RENTABILIDADE E LTV")
print("=" * 100)

perfil_analise = df_clean.groupby('Perfil').agg({
    'Nome da Escola': 'count',
    'Faturamento_Numerico': ['sum', 'mean', 'median'],
    'Alunado': ['sum', 'mean']
}).round(2)

perfil_analise.columns = ['Qtd Escolas', 'Fat. Total', 'Fat. Médio', 'Fat. Mediano', 'Alunos Total', 'Alunos Médio']

print("\n--- Análise por Perfil de Escola ---")
print(perfil_analise.to_string())

print("\n--- Ranking por Faturamento Médio (proxy para LTV/Rentabilidade) ---")
ranking = perfil_analise.sort_values('Fat. Médio', ascending=False)
for idx, (perfil, dados) in enumerate(ranking.iterrows(), 1):
    print(f"{idx}. {perfil}: R$ {dados['Fat. Médio']:,.2f} (média) | {int(dados['Qtd Escolas'])} escolas")

# ANÁLISE 8: UPSELL
print("\n" + "=" * 100)
print("8. ESCOLAS PEDAGÓGICAS COM POTENCIAL DE UPSELL - LISTA DE PRIORIDADES")
print("=" * 100)

# Filtrar escolas pedagógicas
pedagogicas = df_clean[df_clean['Perfil'] == 'Pedagógica'].copy()

# Critérios de priorização
pedagogicas['Score_Upsell'] = 0

# Pontuação por segmentos parciais
pedagogicas.loc[pedagogicas['Cobertura_Segmentos'] == 'Parcial', 'Score_Upsell'] += 3

# Pontuação por horas baixas
pedagogicas.loc[pedagogicas['Horas_Numerico'] <= 3, 'Score_Upsell'] += 2

# Pontuação por alunado alto
pedagogicas.loc[pedagogicas['Alunado'] > 500, 'Score_Upsell'] += 3
pedagogicas.loc[(pedagogicas['Alunado'] > 300) & (pedagogicas['Alunado'] <= 500), 'Score_Upsell'] += 2

# Ordenar por score
pedagogicas_sorted = pedagogicas.sort_values('Score_Upsell', ascending=False)

print("\n--- LISTA PRIORIZADA DE UPSELL (Escolas Pedagógicas) ---\n")
for idx, (_, row) in enumerate(pedagogicas_sorted.iterrows(), 1):
    print(f"{idx}. {row['Nome da Escola']} [Score: {row['Score_Upsell']}]")
    print(f"   • Alunado: {int(row['Alunado']) if pd.notna(row['Alunado']) else 'N/D'} alunos")
    print(f"   • Faturamento Atual: R$ {row['Faturamento_Numerico']:,.2f}")
    print(f"   • Horas: {row['Horas ']} ({row['Horas_Numerico']}h)")
    print(f"   • Segmentos: {row['Seg']} [{row['Cobertura_Segmentos']}]")
    print(f"   • Tipo: {row['Tipo_Campo']} (cidade {'com' if row['Tipo_Campo'] == 'Brownfield' else 'sem'} franqueado)")
    print(f"   • Oportunidades:")
    
    oportunidades = []
    if row['Cobertura_Segmentos'] == 'Parcial':
        oportunidades.append("Expandir para mais segmentos educacionais")
    if pd.notna(row['Horas_Numerico']) and row['Horas_Numerico'] <= 3:
        oportunidades.append(f"Aumentar carga horária (de {row['Horas_Numerico']} para 4-5h)")
    if row['Alunado'] > 500:
        oportunidades.append("Alto volume de alunos = alto potencial de receita adicional")
    
    for op in oportunidades:
        print(f"     → {op}")
    print()

# RESUMO EXECUTIVO
print("\n" + "=" * 100)
print("RESUMO EXECUTIVO - PRINCIPAIS INSIGHTS (VERSÃO CORRIGIDA)")
print("=" * 100)
print(f"""
📊 VISÃO GERAL:
   • Total de Escolas: {len(df_clean)}
   • Faturamento Total: R$ {df_clean['Faturamento_Numerico'].sum():,.2f}
   • Ticket Médio: R$ {df_clean['Faturamento_Numerico'].mean():,.2f}
   • Alunos Totais: {int(df_clean['Alunado'].sum())}

🏗️ BROWNFIELD vs GREENFIELD (CORRIGIDO):
   • Brownfield (COM franqueado RB): {brownfield_count} escolas ({brownfield_perc:.1f}%) - R$ {brownfield_fat:,.2f}
   • Greenfield (SEM franqueado RB): {greenfield_count} escolas ({greenfield_perc:.1f}%) - R$ {greenfield_fat:,.2f}

📚 INGLÊS ANTERIOR:
   • Tinham inglês anterior: {tinha_ingles} escolas ({tinha_ingles/len(df_clean)*100:.1f}%)
   • Não tinham ou insatisfeitas: {nao_tinha} escolas ({nao_tinha/len(df_clean)*100:.1f}%)

💰 PERFIL MAIS RENTÁVEL:
   • {ranking.index[0]}: R$ {ranking.iloc[0]['Fat. Médio']:,.2f} (faturamento médio)

🎓 TOP OPORTUNIDADES DE UPSELL:
""")

for idx, (_, row) in enumerate(pedagogicas_sorted.head(3).iterrows(), 1):
    print(f"   {idx}. {row['Nome da Escola']} - Score: {row['Score_Upsell']} [{row['Tipo_Campo']}]")

print("\n" + "=" * 100)
