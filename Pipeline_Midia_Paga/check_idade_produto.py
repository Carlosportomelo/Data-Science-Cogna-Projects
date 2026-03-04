import pandas as pd
import re

print("Lendo dados...")
df = pd.read_excel(r'outputs\visão resumida\hubspot_resumido_Dash.xlsx', 
                    sheet_name='Visao_Granular_Final',
                    usecols=['Data', 'Tipo'],
                    engine='openpyxl')

# Conversão de data
df['Data'] = pd.to_datetime(df['Data']).dt.normalize()

# Filtrar período OUT/2025 A MAR/2026
periodo = df[(df['Data'] >= '2025-10-01') & 
             (df['Data'] <= '2026-03-31')].copy()

print(f"Total de leads no período (out 2025 - mar 2026): {len(periodo)}\n")

# Verificar se há idade no formato: "Nome - X - Local" onde X é um número
def tem_idade_no_tipo(tipo_str):
    if pd.isna(tipo_str):
        return False
    # Procurar por padrão: - número - (idade entre hífens)
    # Ou número sozinho no meio do texto
    padrao = r'-\s*(\d+)\s*-'
    match = re.search(padrao, str(tipo_str))
    return match is not None

periodo['Tem_Idade'] = periodo['Tipo'].apply(tem_idade_no_tipo)

total_leads = len(periodo)
com_idade = periodo['Tem_Idade'].sum()
sem_idade = total_leads - com_idade
percentual_sem_idade = (sem_idade / total_leads) * 100

print(f"Leads COM idade no nome do produto: {com_idade:,} ({(com_idade/total_leads)*100:.2f}%)")
print(f"Leads SEM idade no nome do produto: {sem_idade:,} ({percentual_sem_idade:.2f}%)")
print(f"\nTotal: {total_leads:,}")

# Mostrar alguns exemplos de cada tipo
print("\n" + "="*70)
print("EXEMPLOS COM IDADE:")
print("="*70)
exemplos_com = periodo[periodo['Tem_Idade']]['Tipo'].head(10)
for ex in exemplos_com:
    print(f"  {ex}")

print("\n" + "="*70)
print("EXEMPLOS SEM IDADE:")
print("="*70)
exemplos_sem = periodo[~periodo['Tem_Idade']]['Tipo'].head(10)
for ex in exemplos_sem:
    print(f"  {ex}")
