import pandas as pd
import numpy as np

print("Carregando bases...")
df_hoje = pd.read_csv('contagem-negocios.csv')
df_sexta = pd.read_csv('hubspot_leads.csv')

# Identificar leads que sumiram
ids_hoje = set(df_hoje['Record ID'])
ids_sexta = set(df_sexta['Record ID'])
leads_sumidos_ids = ids_sexta - ids_hoje

df_sumidos = df_sexta[df_sexta['Record ID'].isin(leads_sumidos_ids)].copy()

print(f"Total de leads que sumiram: {len(df_sumidos):,}\n")

print("="*100)
print("🔍 ANÁLISE DE DUPLICAÇÃO - LEADS QUE SUMIRAM")
print("="*100)

# Limpar e normalizar nomes para melhor comparação
def limpar_nome(nome):
    if pd.isna(nome):
        return ""
    nome = str(nome).strip().upper()
    # Remover sufixos comuns que indicam duplicação
    nome = nome.split(' - ')[0] if ' - ' in nome else nome
    return nome

df_sumidos['Nome Limpo'] = df_sumidos['Nome do negócio'].apply(limpar_nome)
df_hoje['Nome Limpo'] = df_hoje['Nome do negócio'].apply(limpar_nome)
df_sexta['Nome Limpo'] = df_sexta['Nome do negócio'].apply(limpar_nome)

# 1. Verificar duplicações DENTRO dos leads que sumiram
print("\n📋 PARTE 1: NOMES DUPLICADOS DENTRO DOS LEADS QUE SUMIRAM")
print("-" * 100)

duplicados_sumidos = df_sumidos[df_sumidos['Nome Limpo'] != ''].groupby('Nome Limpo').agg({
    'Record ID': ['count', lambda x: ', '.join(map(str, x))],
    'Unidade Desejada': lambda x: ', '.join(set(x.astype(str))),
    'Etapa do negócio': lambda x: ', '.join(set(x)),
    'Data de criação': lambda x: ', '.join(x.astype(str))
}).reset_index()

duplicados_sumidos.columns = ['Nome', 'Qtd IDs', 'Record IDs', 'Unidades', 'Etapas', 'Datas Criação']
duplicados_sumidos = duplicados_sumidos[duplicados_sumidos['Qtd IDs'] > 1].sort_values('Qtd IDs', ascending=False)

print(f"\nTotal de nomes com múltiplos Record IDs que sumiram: {len(duplicados_sumidos)}")
print(f"Total de leads duplicados que sumiram: {duplicados_sumidos['Qtd IDs'].sum()}")

if len(duplicados_sumidos) > 0:
    print(f"\nTOP 20 NOMES MAIS DUPLICADOS QUE SUMIRAM:")
    print("-" * 100)
    for idx, row in duplicados_sumidos.head(20).iterrows():
        print(f"\n{row['Nome']}")
        print(f"  • Quantidade de IDs: {row['Qtd IDs']}")
        print(f"  • Record IDs: {row['Record IDs'][:100]}...")
        print(f"  • Unidades: {row['Unidades'][:80]}")

# 2. Verificar se os mesmos nomes ainda existem na base de HOJE
print("\n" + "="*100)
print("📋 PARTE 2: LEADS QUE SUMIRAM MAS O MESMO NOME AINDA EXISTE NA BASE DE HOJE")
print("="*100)

nomes_sumidos = set(df_sumidos['Nome Limpo']) - {''}
nomes_hoje = set(df_hoje['Nome Limpo']) - {''}

nomes_que_permaneceram = nomes_sumidos & nomes_hoje

print(f"\nNomes que sumiram: {len(nomes_sumidos):,}")
print(f"Nomes que ainda existem hoje: {len(nomes_que_permaneceram):,}")
print(f"Percentual: {len(nomes_que_permaneceram)/len(nomes_sumidos)*100:.1f}%")

# Análise detalhada dos que permaneceram
casos_duplicacao = []

for nome in nomes_que_permaneceram:
    # Leads que sumiram com esse nome
    leads_sumidos_nome = df_sumidos[df_sumidos['Nome Limpo'] == nome]
    # Leads que ainda existem com esse nome
    leads_hoje_nome = df_hoje[df_hoje['Nome Limpo'] == nome]
    
    casos_duplicacao.append({
        'Nome': nome,
        'IDs Sumidos': len(leads_sumidos_nome),
        'IDs Hoje': len(leads_hoje_nome),
        'Total IDs': len(leads_sumidos_nome) + len(leads_hoje_nome),
        'Record IDs Sumidos': ', '.join(map(str, leads_sumidos_nome['Record ID'].tolist())),
        'Record IDs Hoje': ', '.join(map(str, leads_hoje_nome['Record ID'].tolist())),
        'Etapas Sumidas': ', '.join(set(leads_sumidos_nome['Etapa do negócio'])),
        'Etapas Hoje': ', '.join(set(leads_hoje_nome['Etapa do negócio'])),
        'Unidades Sumidas': ', '.join(set(leads_sumidos_nome['Unidade Desejada'].astype(str))),
        'Unidades Hoje': ', '.join(set(leads_hoje_nome['Unidade Desejada'].astype(str)))
    })

df_casos = pd.DataFrame(casos_duplicacao).sort_values('Total IDs', ascending=False)

print(f"\nTOP 30 CASOS DE POSSÍVEL LIMPEZA DE DUPLICATAS:")
print("-" * 100)
print(f"{'Nome':<40} {'IDs Sumidos':<12} {'IDs Hoje':<10} {'Total':<8}")
print("-" * 100)

for idx, row in df_casos.head(30).iterrows():
    print(f"{row['Nome'][:38]:<40} {row['IDs Sumidos']:<12} {row['IDs Hoje']:<10} {row['Total IDs']:<8}")

# 3. Análise COMPLETA na base de sexta-feira
print("\n" + "="*100)
print("📋 PARTE 3: DUPLICAÇÕES NA BASE DE SEXTA-FEIRA (ANTES DA LIMPEZA)")
print("="*100)

duplicados_sexta = df_sexta[df_sexta['Nome Limpo'] != ''].groupby('Nome Limpo').agg({
    'Record ID': 'count'
}).reset_index()
duplicados_sexta.columns = ['Nome', 'Qtd IDs']
duplicados_sexta = duplicados_sexta[duplicados_sexta['Qtd IDs'] > 1].sort_values('Qtd IDs', ascending=False)

print(f"\nTotal de nomes duplicados na base de sexta: {len(duplicados_sexta):,}")
print(f"Total de leads com nomes duplicados: {duplicados_sexta['Qtd IDs'].sum():,}")

# 4. Análise COMPLETA na base de hoje
print("\n" + "="*100)
print("📋 PARTE 4: DUPLICAÇÕES NA BASE DE HOJE (DEPOIS DA LIMPEZA)")
print("="*100)

duplicados_hoje = df_hoje[df_hoje['Nome Limpo'] != ''].groupby('Nome Limpo').agg({
    'Record ID': 'count'
}).reset_index()
duplicados_hoje.columns = ['Nome', 'Qtd IDs']
duplicados_hoje = duplicados_hoje[duplicados_hoje['Qtd IDs'] > 1].sort_values('Qtd IDs', ascending=False)

print(f"\nTotal de nomes duplicados na base de hoje: {len(duplicados_hoje):,}")
print(f"Total de leads com nomes duplicados: {duplicados_hoje['Qtd IDs'].sum():,}")

# Comparação
reducao = len(duplicados_sexta) - len(duplicados_hoje)
print(f"\n✅ Redução de nomes duplicados: {reducao:,} ({reducao/len(duplicados_sexta)*100:.1f}%)")

# 5. RESUMO EXECUTIVO
print("\n" + "="*100)
print("📊 RESUMO EXECUTIVO - HIPÓTESE DE LIMPEZA DE DUPLICATAS")
print("="*100)

leads_duplicados_sumidos = len(df_casos)
percentual_duplicados = (leads_duplicados_sumidos / len(df_sumidos)) * 100

print(f"""
LEADS QUE SUMIRAM: {len(df_sumidos):,}
  • Nomes únicos: {len(nomes_sumidos):,}
  • Nomes que ainda existem hoje: {len(nomes_que_permaneceram):,} ({len(nomes_que_permaneceram)/len(nomes_sumidos)*100:.1f}%)
  • Possíveis duplicatas removidas: {leads_duplicados_sumidos:,} ({percentual_duplicados:.1f}%)

DUPLICAÇÃO GERAL:
  • Sexta-feira - Nomes duplicados: {len(duplicados_sexta):,}
  • Hoje - Nomes duplicados: {len(duplicados_hoje):,}
  • Redução de duplicatas: {reducao:,} nomes

CONCLUSÃO:
""")

if percentual_duplicados > 50:
    print("✅ MAIS DE 50% dos leads que sumiram têm o mesmo nome em registros que ainda existem.")
    print("   Isso CONFIRMA que houve uma limpeza de duplicatas no HubSpot.")
else:
    print(f"⚠️  Apenas {percentual_duplicados:.1f}% dos leads que sumiram são duplicatas.")
    print("   A limpeza de duplicatas não explica completamente os 1.063 leads removidos.")

# Salvar análise detalhada em Excel
print("\n" + "="*100)
print("💾 GERANDO ARQUIVO EXCEL COM ANÁLISE DE DUPLICATAS...")
print("="*100)

with pd.ExcelWriter('ANALISE_DUPLICATAS_LEADS_SUMIDOS.xlsx', engine='openpyxl') as writer:
    # Casos de duplicação
    df_casos.to_excel(writer, sheet_name='Duplicatas Identificadas', index=False)
    
    # Duplicados que sumiram
    if len(duplicados_sumidos) > 0:
        duplicados_sumidos.to_excel(writer, sheet_name='Duplicados Sumidos', index=False)
    
    # Top duplicados sexta
    duplicados_sexta.head(100).to_excel(writer, sheet_name='Top Duplicados Sexta', index=False)
    
    # Top duplicados hoje
    duplicados_hoje.head(100).to_excel(writer, sheet_name='Top Duplicados Hoje', index=False)
    
    # Resumo
    resumo_data = {
        'Métrica': [
            'Total de leads que sumiram',
            'Nomes únicos que sumiram',
            'Nomes que ainda existem hoje',
            'Percentual de duplicatas',
            '',
            'Nomes duplicados (Sexta)',
            'Nomes duplicados (Hoje)',
            'Redução de duplicatas'
        ],
        'Valor': [
            len(df_sumidos),
            len(nomes_sumidos),
            len(nomes_que_permaneceram),
            f"{percentual_duplicados:.1f}%",
            '',
            len(duplicados_sexta),
            len(duplicados_hoje),
            reducao
        ]
    }
    pd.DataFrame(resumo_data).to_excel(writer, sheet_name='Resumo', index=False)

print("\n✅ ARQUIVO CRIADO: ANALISE_DUPLICATAS_LEADS_SUMIDOS.xlsx")
print("\n📑 Abas criadas:")
print("   • Duplicatas Identificadas - Nomes que sumiram mas ainda existem")
print("   • Duplicados Sumidos - Múltiplos IDs para o mesmo nome que sumiram")
print("   • Top Duplicados Sexta - Maiores duplicatas antes da limpeza")
print("   • Top Duplicados Hoje - Maiores duplicatas após limpeza")
print("   • Resumo - Estatísticas consolidadas")
print("\n" + "="*100)
