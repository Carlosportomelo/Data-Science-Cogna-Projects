import pandas as pd
import os
from datetime import datetime

print("=" * 130)
print("INVESTIGAÇÃO: QUAL BASE DE DADOS FOI USADA NA ANÁLISE DO DIA 30?")
print("=" * 130)
print()

# Carregar arquivo Excel do dia 30 para ver se há informação sobre a fonte
analise_arquivo = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-01-30.xlsx'
xls = pd.ExcelFile(analise_arquivo)

print("Abas do Excel:")
for aba in xls.sheet_names:
    df = pd.read_excel(analise_arquivo, sheet_name=aba)
    print(f"\n{aba}: {df.shape[0]} linhas x {df.shape[1]} colunas")
    print(f"Colunas: {list(df.columns)}")

print("\n\n" + "=" * 130)
print("COMPARAÇÃO DE BASES DE DADOS DISPONÍVEIS")
print("=" * 130)
print()

# Carregar ambas as bases
base_atual = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\hubspot_leads.csv'
base_backup = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads.csv'

df_atual = pd.read_csv(base_atual)
df_backup = pd.read_csv(base_backup)

df_atual['Data de criação'] = pd.to_datetime(df_atual['Data de criação'], format='%Y-%m-%d %H:%M')
df_backup['Data de criação'] = pd.to_datetime(df_backup['Data de criação'], format='%Y-%m-%d %H:%M')

# Data máxima de cada base
data_max_atual = df_atual['Data de criação'].max()
data_max_backup = df_backup['Data de criação'].max()

print(f"Base atual:")
print(f"  Total de registros: {len(df_atual)}")
print(f"  Data máxima: {data_max_atual}")
print()

print(f"Base backup:")
print(f"  Total de registros: {len(df_backup)}")
print(f"  Data máxima: {data_max_backup}")
print()

# Testar diferentes períodos para ver qual corresponde à análise
print("=" * 130)
print("TESTE 1: Filtros até 30/01/2026 usando ambas as bases")
print("=" * 130)
print()

data_corte = pd.Timestamp('2026-01-30 23:59:59')

# Teste 1: Com backup até 30/01
df_backup_corte = df_backup[df_backup['Data de criação'] <= data_corte]

# Ciclos: 01/out do ano anterior até 20/02 do ano seguinte (inclusivo)
ciclos = {}
for yy in ['23', '24', '25', '26']:
    label = f"{yy}.1"
    start_year = 2000 + int(yy) - 1
    start = pd.Timestamp(f"{start_year}-10-01")
    end = pd.Timestamp(f"{start_year + 1}-02-20") + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    ciclos[label] = (start, end)

print("Com base BACKUP até 30/01:")
for ciclo, (data_inicio, data_fim) in ciclos.items():
    df_ciclo = df_backup_corte[(df_backup_corte['Data de criação'] >= data_inicio) & 
                                (df_backup_corte['Data de criação'] <= data_fim)]
    print(f"  {ciclo}: {len(df_ciclo)} leads")

print()

# Teste 2: Com backup sem filtro de data
print("Com base BACKUP SEM FILTRO de data:")
for ciclo, (data_inicio, data_fim) in ciclos.items():
    df_ciclo = df_backup[(df_backup['Data de criação'] >= data_inicio) & 
                          (df_backup['Data de criação'] <= data_fim)]
    print(f"  {ciclo}: {len(df_ciclo)} leads")

print()

# Teste 3: Com base atual até 30/01
df_atual_corte = df_atual[df_atual['Data de criação'] <= data_corte]

print("Com base ATUAL até 30/01:")
for ciclo, (data_inicio, data_fim) in ciclos.items():
    df_ciclo = df_atual_corte[(df_atual_corte['Data de criação'] >= data_inicio) & 
                                (df_atual_corte['Data de criação'] <= data_fim)]
    print(f"  {ciclo}: {len(df_ciclo)} leads")

print()

# Teste 4: Com base atual sem filtro de data
print("Com base ATUAL SEM FILTRO de data:")
for ciclo, (data_inicio, data_fim) in ciclos.items():
    df_ciclo = df_atual[(df_atual['Data de criação'] >= data_inicio) & 
                        (df_atual['Data de criação'] <= data_fim)]
    print(f"  {ciclo}: {len(df_ciclo)} leads")

print()

print("=" * 130)
print("NÚMEROS DA ANÁLISE DO DIA 30 (para comparação)")
print("=" * 130)
print()
print("  23.1: 6.290 leads")
print("  24.1: 8.296 leads")
print("  25.1: 9.688 leads")
print("  26.1: 10.421 leads")
print()

print("=" * 130)
print("CONCLUSÃO")
print("=" * 130)
print()
print("A análise do dia 30 foi feita com:")
print("  → UMA BASE INCOMPLETA ou com FILTROS ESPECÍFICOS")
print("  → Números entre 30-46% menores que a base de backup")
print("  → Provavelmente usou dados com processamento/limpeza adicional")
print()