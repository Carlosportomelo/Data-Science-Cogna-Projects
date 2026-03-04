import pandas as pd
from datetime import datetime

df = pd.read_csv(r'C:\Users\a483650\Projetos\06_Analise_Funil_RedBalloon\data\hubspot_leads_redballoon.csv')

# Converter datas
df['Data de criação'] = pd.to_datetime(df['Data de criação'])
df['Data da última atividade'] = pd.to_datetime(df['Data da última atividade'], errors='coerce')

# Definir ciclos
def definir_ciclo(data):
    if pd.isna(data):
        return None
    ano = data.year
    mes = data.month
    if mes >= 10:
        return f"{str(ano + 1)[-2:]}.1"
    elif mes <= 3:
        return f"{str(ano)[-2:]}.1"
    else:
        return None

df['ciclo'] = df['Data de criação'].apply(definir_ciclo)

# Filtrar ciclo 26.1
df_ciclo = df[df['ciclo'] == '26.1'].copy()

# Calcular idade e recência
hoje = datetime.now()
df_ciclo['idade_dias'] = (hoje - df_ciclo['Data de criação']).dt.days
df_ciclo['dias_desde_ultima_atividade'] = (hoje - df_ciclo['Data da última atividade']).dt.days

# Filtrar leads D-1 (idade ≤1 dia)
leads_d1 = df_ciclo[df_ciclo['idade_dias'] <= 1].copy()

print("=" * 80)
print("ANÁLISE DE LEADS D-1 (IDADE ≤1 DIA)")
print("=" * 80)
print(f"\nTotal de leads D-1: {len(leads_d1)}")

# Quantos têm atividades?
leads_d1_com_atividade = leads_d1[leads_d1['Número de atividades de vendas'] > 0]
print(f"Leads D-1 com atividades (>0): {len(leads_d1_com_atividade)}")

# Quantos têm data da última atividade preenchida?
leads_d1_com_data = leads_d1[leads_d1['Data da última atividade'].notna()]
print(f"Leads D-1 com 'Data da última atividade' preenchida: {len(leads_d1_com_data)}")

# Quantos têm atividade D-1 (recência ≤1 dia)?
leads_d1_ativ_d1 = leads_d1[leads_d1['dias_desde_ultima_atividade'] <= 1]
print(f"Leads D-1 com atividade D-1 (recência ≤1 dia): {len(leads_d1_ativ_d1)}")

# Ver exemplos de inconsistências
print("\n" + "=" * 80)
print("LEADS D-1 COM ATIVIDADES MAS SEM DATA DA ÚLTIMA ATIVIDADE:")
print("=" * 80)

inconsistentes = leads_d1[
    (leads_d1['Número de atividades de vendas'] > 0) & 
    (leads_d1['Data da última atividade'].isna())
]

if len(inconsistentes) > 0:
    print(f"\nEncontrados {len(inconsistentes)} leads inconsistentes!")
    print("\nExemplos:")
    print(inconsistentes[['Record ID', 'Nome do negócio', 'Data de criação', 
                          'Número de atividades de vendas', 'Data da última atividade']].head(10).to_string())
else:
    print("\nNenhum lead inconsistente encontrado.")

# Ver leads com atividade mas recência > D-1
print("\n" + "=" * 80)
print("LEADS D-1 COM ATIVIDADES MAS ÚLTIMA ATIVIDADE > D-1:")
print("=" * 80)

estranhos = leads_d1[
    (leads_d1['Número de atividades de vendas'] > 0) & 
    (leads_d1['Data da última atividade'].notna()) &
    (leads_d1['dias_desde_ultima_atividade'] > 1)
]

if len(estranhos) > 0:
    print(f"\nEncontrados {len(estranhos)} leads estranhos!")
    print("(Leads criados há ≤1 dia mas última atividade foi há >1 dia)")
    print("\nExemplos:")
    print(estranhos[['Record ID', 'Nome do negócio', 'Data de criação', 
                     'Data da última atividade', 'idade_dias', 
                     'dias_desde_ultima_atividade']].head(10).to_string())
else:
    print("\nNenhum lead estranho encontrado.")

print("\n" + "=" * 80)
print("RESUMO:")
print("=" * 80)
print(f"Leads D-1: {len(leads_d1)}")
print(f"  Com atividades: {len(leads_d1_com_atividade)} ({100*len(leads_d1_com_atividade)/len(leads_d1):.1f}%)")
print(f"  Com data preenchida: {len(leads_d1_com_data)} ({100*len(leads_d1_com_data)/len(leads_d1):.1f}%)")
print(f"  Com atividade D-1: {len(leads_d1_ativ_d1)} ({100*len(leads_d1_ativ_d1)/len(leads_d1):.1f}%)")
print(f"\nDiferença: {len(leads_d1_com_atividade) - len(leads_d1_ativ_d1)} leads com atividades MAS sem data (ou data antiga)")
