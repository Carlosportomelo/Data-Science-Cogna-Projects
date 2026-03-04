import pandas as pd
import numpy as np

# Carregar os arquivos
print("Carregando arquivos...")
df_hoje = pd.read_csv('contagem-negocios.csv')
df_sexta = pd.read_csv('hubspot_leads.csv')

print(f"Leads hoje: {len(df_hoje):,}")
print(f"Leads sexta-feira: {len(df_sexta):,}")
print(f"Diferença: {len(df_sexta) - len(df_hoje):,} leads\n")

# Criar índices por Record ID para comparação rápida
df_hoje_idx = df_hoje.set_index('Record ID')
df_sexta_idx = df_sexta.set_index('Record ID')

# Identificar leads que sumiram
ids_hoje = set(df_hoje['Record ID'])
ids_sexta = set(df_sexta['Record ID'])

leads_sumidos_ids = ids_sexta - ids_hoje
leads_novos_ids = ids_hoje - ids_sexta

print(f"✅ Leads novos (não estavam sexta): {len(leads_novos_ids):,}")
print(f"❌ Leads que REALMENTE SUMIRAM: {len(leads_sumidos_ids):,}\n")

# Filtrar leads que realmente sumiram
df_sumidos = df_sexta[df_sexta['Record ID'].isin(leads_sumidos_ids)].copy()

print("="*80)
print("ANÁLISE DETALHADA - LEADS QUE REALMENTE DESAPARECERAM")
print("="*80)

# Análise por etapa
print("\n📊 DISTRIBUIÇÃO POR ETAPA DO NEGÓCIO:")
etapas = df_sumidos['Etapa do negócio'].value_counts()
print(etapas)

print("\n📊 TOP 15 PROPRIETÁRIOS MAIS AFETADOS:")
proprietarios = df_sumidos['Proprietário do negócio'].value_counts().head(15)
for prop, count in proprietarios.items():
    print(f"{prop:50s}: {count:4d} leads")

print("\n📊 TOP 15 UNIDADES MAIS AFETADAS:")
unidades = df_sumidos['Unidade Desejada'].value_counts().head(15)
for unid, count in unidades.items():
    print(f"{unid:30s}: {count:4d} leads")

print("\n📊 LEADS SUMIDOS COM MATRÍCULA CONCLUÍDA POR PROPRIETÁRIO:")
matriculas = df_sumidos[df_sumidos['Etapa do negócio'] == 'MATRÍCULA CONCLUÍDA (Red Balloon - Unidades de Rua)']
print(f"\nTotal de matrículas concluídas que sumiram: {len(matriculas)}")
print("\nTop proprietários:")
print(matriculas['Proprietário do negócio'].value_counts().head(10))

print("\n📅 DISTRIBUIÇÃO POR DATA DE CRIAÇÃO:")
df_sumidos['Data de criação'] = pd.to_datetime(df_sumidos['Data de criação'])
df_sumidos['Ano-Mês'] = df_sumidos['Data de criação'].dt.to_period('M')
datas = df_sumidos['Ano-Mês'].value_counts().sort_index()
print(datas)

# Análise de fonte de tráfego
print("\n📊 FONTE ORIGINAL DO TRÁFEGO:")
fonte_counts = df_sumidos['Fonte original do tráfego'].fillna('(vazio)').value_counts()
print(fonte_counts)

# Verificar se há padrão nos Record IDs
print("\n📊 ANÁLISE DOS RECORD IDs:")
record_ids = list(leads_sumidos_ids)
if len(record_ids) > 0:
    record_ids_num = sorted([int(x) for x in record_ids])
    print(f"Menor Record ID: {record_ids_num[0]:,}")
    print(f"Maior Record ID: {record_ids_num[-1]:,}")
    print(f"Média Record ID: {np.mean(record_ids_num):,.0f}")
    
    # Verificar se são IDs consecutivos ou com padrão
    print(f"\nPrimeiros 20 Record IDs que sumiram:")
    for i in range(min(20, len(record_ids_num))):
        rid = record_ids_num[i]
        # Buscar no dataframe original já que pode ter conversão de tipo
        lead_info = df_sumidos[df_sumidos['Record ID'].astype(str) == str(rid)].iloc[0]
        print(f"  {rid:11,} | {str(lead_info['Nome do negócio'])[:30]:30s} | {str(lead_info['Etapa do negócio'])[:50]:50s}")

# Verificar padrões específicos
print("\n" + "="*80)
print("ANÁLISE DE PADRÕES CRÍTICOS")
print("="*80)

# Leads de matrícula concluída que sumiram
print("\n🚨 MATRÍCULAS CONCLUÍDAS QUE DESAPARECERAM (Primeiros 20):")
matriculas_sorted = matriculas.sort_values('Data de criação', ascending=False)
colunas = ['Record ID', 'Nome do negócio', 'Data de criação', 'Unidade Desejada', 'Proprietário do negócio', 'Data de fechamento']
print(matriculas_sorted[colunas].head(20).to_string(index=False, max_colwidth=30))

# Leads perdidos que sumiram
print("\n📉 NEGÓCIOS PERDIDOS QUE DESAPARECERAM:")
perdidos = df_sumidos[df_sumidos['Etapa do negócio'] == 'NEGÓCIO PERDIDO (Red Balloon - Unidades de Rua)']
print(f"Total: {len(perdidos)}")
print("\nTop proprietários:")
print(perdidos['Proprietário do negócio'].value_counts().head(10))

# Novos negócios que sumiram
print("\n📝 NOVOS NEGÓCIOS QUE DESAPARECERAM:")
novos = df_sumidos[df_sumidos['Etapa do negócio'] == 'NOVO NEGÓCIO (Red Balloon - Unidades de Rua)']
print(f"Total: {len(novos)}")
print("\nDistribuição por data de criação:")
novos['Data criação'] = novos['Data de criação'].dt.date
print(novos['Data criação'].value_counts().sort_index().tail(10))

# Salvar relatórios detalhados
print("\n" + "="*80)
print("GERANDO ARQUIVOS DE SAÍDA")
print("="*80)

# Arquivo com todos os leads sumidos
df_sumidos_sorted = df_sumidos.sort_values('Data de criação', ascending=False)
df_sumidos_sorted.to_csv('leads_sumidos_completo.csv', index=False)
print(f"✅ leads_sumidos_completo.csv - {len(df_sumidos):,} leads")

# Arquivo só com matrículas concluídas
matriculas_sorted.to_csv('matriculas_sumidas.csv', index=False)
print(f"✅ matriculas_sumidas.csv - {len(matriculas):,} leads")

# Arquivo só com negócios perdidos
perdidos.to_csv('negocios_perdidos_sumidos.csv', index=False)
print(f"✅ negocios_perdidos_sumidos.csv - {len(perdidos):,} leads")

# Arquivo só com novos negócios
novos.to_csv('novos_negocios_sumidos.csv', index=False)
print(f"✅ novos_negocios_sumidos.csv - {len(novos):,} leads")

print("\n" + "="*80)
print("RESUMO EXECUTIVO")
print("="*80)
print(f"""
❌ TOTAL DE LEADS QUE DESAPARECERAM: {len(leads_sumidos_ids):,}

Por Etapa:
  • Matrículas Concluídas: {len(matriculas):,} ({len(matriculas)/len(df_sumidos)*100:.1f}%)
  • Negócios Perdidos: {len(perdidos):,} ({len(perdidos)/len(df_sumidos)*100:.1f}%)
  • Novos Negócios: {len(novos):,} ({len(novos)/len(df_sumidos)*100:.1f}%)
  • Outros: {len(df_sumidos) - len(matriculas) - len(perdidos) - len(novos):,}

🚨 ALERTA: Matrículas concluídas NÃO deveriam desaparecer do histórico!
""")
