import pandas as pd
import re

print("Lendo HubSpot original...")
df = pd.read_csv(r'C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_atual.csv')

# Converter data
df['Data de criação'] = pd.to_datetime(df['Data de criação'])

# Filtrar período OUT/2025 A MAR/2026
periodo = df[(df['Data de criação'] >= '2025-10-01') & 
             (df['Data de criação'] <= '2026-03-31')].copy()

print(f"Total de leads no período (out 2025 - mar 2026): {len(periodo)}\n")

# Verificar se há idade no "Nome do negócio"
# Formato esperado: "Nome - X - Local" onde X é a idade (número)
def tem_idade_no_nome(nome_str):
    if pd.isna(nome_str):
        return False
    # Padrão: - número - (idade entre hífens)
    padrao = r'-\s*(\d+)\s*-'
    match = re.search(padrao, str(nome_str))
    return match is not None

periodo['Tem_Idade'] = periodo['Nome do negócio'].apply(tem_idade_no_nome)

total_leads = len(periodo)
com_idade = periodo['Tem_Idade'].sum()
sem_idade = total_leads - com_idade
percentual_com_idade = (com_idade / total_leads) * 100
percentual_sem_idade = (sem_idade / total_leads) * 100

print("="*70)
print(f"Leads COM idade no nome: {com_idade:,} ({percentual_com_idade:.2f}%)")
print(f"Leads SEM idade no nome: {sem_idade:,} ({percentual_sem_idade:.2f}%)")
print(f"Total: {total_leads:,}")
print("="*70)

# Mostrar alguns exemplos de cada tipo
print("\nEXEMPLOS COM IDADE:")
print("-"*70)
exemplos_com = periodo[periodo['Tem_Idade']]['Nome do negócio'].head(15)
for ex in exemplos_com:
    print(f"  {ex}")

print("\n\nEXEMPLOS SEM IDADE:")
print("-"*70)
exemplos_sem = periodo[~periodo['Tem_Idade']]['Nome do negócio'].head(15)
for ex in exemplos_sem:
    print(f"  {ex}")
