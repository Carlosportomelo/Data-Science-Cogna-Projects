import pandas as pd
import numpy as np

# Carregar os arquivos
print("Carregando arquivos...")
df_hoje = pd.read_csv('contagem-negocios.csv')
df_sexta = pd.read_csv('hubspot_leads.csv')

print(f"Leads hoje: {len(df_hoje):,}")
print(f"Leads sexta-feira: {len(df_sexta):,}")
print(f"Diferença: {len(df_sexta) - len(df_hoje):,} leads\n")

# Identificar leads que sumiram
ids_hoje = set(df_hoje['Record ID'])
ids_sexta = set(df_sexta['Record ID'])

leads_sumidos = ids_sexta - ids_hoje
leads_novos = ids_hoje - ids_sexta

print(f"Leads que sumiram: {len(leads_sumidos):,}")
print(f"Leads novos (não estavam sexta): {len(leads_novos):,}\n")

# Filtrar apenas os leads que sumiram
df_sumidos = df_sexta[df_sexta['Record ID'].isin(leads_sumidos)].copy()

print("="*80)
print("ANÁLISE DOS LEADS QUE SUMIRAM")
print("="*80)

# Análise por campos
print("\n📊 DISTRIBUIÇÃO POR ETAPA DO NEGÓCIO:")
print(df_sumidos['Etapa do negócio'].value_counts())

print("\n📊 DISTRIBUIÇÃO POR PIPELINE:")
print(df_sumidos['Pipeline'].value_counts())

print("\n📊 DISTRIBUIÇÃO POR FONTE ORIGINAL DO TRÁFEGO:")
print(df_sumidos['Fonte original do tráfego'].value_counts())

print("\n📊 DISTRIBUIÇÃO POR PROPRIETÁRIO DO NEGÓCIO:")
top_proprietarios = df_sumidos['Proprietário do negócio'].value_counts().head(10)
print(top_proprietarios)

print("\n📊 DISTRIBUIÇÃO POR UNIDADE DESEJADA:")
top_unidades = df_sumidos['Unidade Desejada'].value_counts().head(10)
print(top_unidades)

print("\n📊 DISTRIBUIÇÃO POR DATA DE CRIAÇÃO:")
df_sumidos['Data de criação'] = pd.to_datetime(df_sumidos['Data de criação'])
df_sumidos['Data criação'] = df_sumidos['Data de criação'].dt.date
print(df_sumidos['Data criação'].value_counts().sort_index().tail(10))

# Análise de padrões específicos
print("\n" + "="*80)
print("ANÁLISE DE PADRÕES ESPECÍFICOS")
print("="*80)

# Verificar se tem algum padrão no Detalhamento da fonte
print("\n📌 Detalhamento da fonte original do tráfego 1:")
print(df_sumidos['Detalhamento da fonte original do tráfego 1'].value_counts().head(10))

print("\n📌 Detalhamento da fonte original do tráfego 2:")
print(df_sumidos['Detalhamento da fonte original do tráfego 2'].value_counts().head(10))

# Verificar leads com valores monetários
df_sumidos['Valor na moeda da empresa'] = pd.to_numeric(df_sumidos['Valor na moeda da empresa'], errors='coerce')
leads_com_valor = df_sumidos[df_sumidos['Valor na moeda da empresa'].notna()]
print(f"\n💰 Leads sumidos com valor monetário: {len(leads_com_valor)}")
if len(leads_com_valor) > 0:
    print(f"Valor total perdido: R$ {leads_com_valor['Valor na moeda da empresa'].sum():,.2f}")

# Verificar leads com Data de fechamento
leads_com_data_fechamento = df_sumidos[df_sumidos['Data de fechamento'].notna() & (df_sumidos['Data de fechamento'] != '')]
print(f"\n📅 Leads sumidos com data de fechamento: {len(leads_com_data_fechamento)}")

# Salvar a lista completa de leads sumidos para análise detalhada
df_sumidos_ordenado = df_sumidos.sort_values('Data de criação', ascending=False)
df_sumidos_ordenado.to_csv('leads_sumidos_analise.csv', index=False)
print(f"\n✅ Arquivo 'leads_sumidos_analise.csv' criado com {len(df_sumidos)} leads para análise detalhada")

# Mostrar alguns exemplos de leads que sumiram
print("\n" + "="*80)
print("EXEMPLOS DE LEADS QUE SUMIRAM (5 mais recentes)")
print("="*80)
colunas_principais = ['Record ID', 'Nome do negócio', 'Data de criação', 'Etapa do negócio', 
                      'Unidade Desejada', 'Proprietário do negócio']
print(df_sumidos_ordenado[colunas_principais].head(10).to_string(index=False))
