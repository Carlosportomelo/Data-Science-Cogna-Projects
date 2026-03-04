import pandas as pd
import numpy as np
from datetime import datetime, timedelta

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
print("🔍 ANÁLISE RIGOROSA DE DUPLICATAS - CRITÉRIOS MÚLTIPLOS")
print("="*100)

# Limpar e normalizar campos
def limpar_texto(texto):
    if pd.isna(texto):
        return ""
    return str(texto).strip().upper()

df_sumidos['Nome Limpo'] = df_sumidos['Nome do negócio'].apply(limpar_texto)
df_sumidos['Unidade Limpa'] = df_sumidos['Unidade Desejada'].apply(limpar_texto)
df_sumidos['Data de criação'] = pd.to_datetime(df_sumidos['Data de criação'])

df_hoje['Nome Limpo'] = df_hoje['Nome do negócio'].apply(limpar_texto)
df_hoje['Unidade Limpa'] = df_hoje['Unidade Desejada'].apply(limpar_texto)
df_hoje['Data de criação'] = pd.to_datetime(df_hoje['Data de criação'])

# =======================================================================================
# CRITÉRIO 1: NOME + UNIDADE (match exato)
# =======================================================================================
print("\n📋 CRITÉRIO 1: DUPLICATAS POR NOME + UNIDADE EXATA")
print("-" * 100)

df_sumidos['Chave_Nome_Unidade'] = df_sumidos['Nome Limpo'] + '|||' + df_sumidos['Unidade Limpa']
df_hoje['Chave_Nome_Unidade'] = df_hoje['Nome Limpo'] + '|||' + df_hoje['Unidade Limpa']

# Filtrar apenas registros com nome e unidade válidos
df_sumidos_valido = df_sumidos[(df_sumidos['Nome Limpo'] != '') & (df_sumidos['Unidade Limpa'] != '')].copy()
df_hoje_valido = df_hoje[(df_hoje['Nome Limpo'] != '') & (df_hoje['Unidade Limpa'] != '')].copy()

chaves_sumidas = set(df_sumidos_valido['Chave_Nome_Unidade'])
chaves_hoje = set(df_hoje_valido['Chave_Nome_Unidade'])

duplicatas_nome_unidade = chaves_sumidas & chaves_hoje

print(f"Leads válidos que sumiram (com nome e unidade): {len(df_sumidos_valido):,}")
print(f"Combinações Nome+Unidade que sumiram: {len(chaves_sumidas):,}")
print(f"Combinações que AINDA EXISTEM hoje: {len(duplicatas_nome_unidade):,}")
print(f"Percentual de possíveis duplicatas: {len(duplicatas_nome_unidade)/len(chaves_sumidas)*100:.1f}%")

# Analisar essas duplicatas em detalhe
casos_nome_unidade = []

for chave in list(duplicatas_nome_unidade)[:500]:  # Limitar para não travar
    nome, unidade = chave.split('|||')
    
    leads_sumidos_chave = df_sumidos_valido[df_sumidos_valido['Chave_Nome_Unidade'] == chave]
    leads_hoje_chave = df_hoje_valido[df_hoje_valido['Chave_Nome_Unidade'] == chave]
    
    casos_nome_unidade.append({
        'Nome': nome[:50],
        'Unidade': unidade[:30],
        'IDs Sumidos': len(leads_sumidos_chave),
        'IDs Hoje': len(leads_hoje_chave),
        'Record IDs Sumidos': ', '.join(map(str, leads_sumidos_chave['Record ID'].tolist()[:5])),
        'Record IDs Hoje': ', '.join(map(str, leads_hoje_chave['Record ID'].tolist()[:5])),
        'Etapas Sumidas': ', '.join(set(leads_sumidos_chave['Etapa do negócio'].head(3))),
        'Etapas Hoje': ', '.join(set(leads_hoje_chave['Etapa do negócio'].head(3))),
        'Proprietários Sumidos': ', '.join(set(leads_sumidos_chave['Proprietário do negócio'].head(3))),
        'Proprietários Hoje': ', '.join(set(leads_hoje_chave['Proprietário do negócio'].head(3)))
    })

df_duplicatas_1 = pd.DataFrame(casos_nome_unidade).sort_values(['IDs Sumidos', 'IDs Hoje'], ascending=False)

print(f"\nTOP 20 DUPLICATAS NOME+UNIDADE (casos com mais registros):")
print("-" * 100)
print(f"{'Nome':<35} {'Unidade':<20} {'IDs Sumidos':<12} {'IDs Hoje':<10}")
print("-" * 100)

for idx, row in df_duplicatas_1.head(20).iterrows():
    print(f"{row['Nome'][:33]:<35} {row['Unidade'][:18]:<20} {row['IDs Sumidos']:<12} {row['IDs Hoje']:<10}")

# =======================================================================================
# CRITÉRIO 2: NOME + UNIDADE + PROPRIETÁRIO (match super exato)
# =======================================================================================
print("\n" + "="*100)
print("📋 CRITÉRIO 2: DUPLICATAS POR NOME + UNIDADE + PROPRIETÁRIO")
print("-" * 100)

df_sumidos_valido['Chave_Completa'] = (df_sumidos_valido['Nome Limpo'] + '|||' + 
                                        df_sumidos_valido['Unidade Limpa'] + '|||' + 
                                        df_sumidos_valido['Proprietário do negócio'].apply(limpar_texto))

df_hoje_valido['Chave_Completa'] = (df_hoje_valido['Nome Limpo'] + '|||' + 
                                     df_hoje_valido['Unidade Limpa'] + '|||' + 
                                     df_hoje_valido['Proprietário do negócio'].apply(limpar_texto))

chaves_completas_sumidas = set(df_sumidos_valido['Chave_Completa'])
chaves_completas_hoje = set(df_hoje_valido['Chave_Completa'])

duplicatas_completas = chaves_completas_sumidas & chaves_completas_hoje

print(f"Combinações Nome+Unidade+Proprietário que AINDA EXISTEM: {len(duplicatas_completas):,}")
print(f"Percentual de duplicatas exatas: {len(duplicatas_completas)/len(chaves_completas_sumidas)*100:.1f}%")

# =======================================================================================
# CRITÉRIO 3: ANALISAR MÚLTIPLOS IDs PARA O MESMO NOME+UNIDADE (dentro da base que sumiu)
# =======================================================================================
print("\n" + "="*100)
print("📋 CRITÉRIO 3: MÚLTIPLOS RECORD IDs PARA O MESMO NOME+UNIDADE (DENTRO DOS SUMIDOS)")
print("-" * 100)

duplicados_internos = df_sumidos_valido.groupby('Chave_Nome_Unidade').agg({
    'Record ID': ['count', lambda x: list(x)],
    'Etapa do negócio': lambda x: list(set(x)),
    'Data de criação': lambda x: list(x),
    'Proprietário do negócio': lambda x: list(set(x))
}).reset_index()

duplicados_internos.columns = ['Chave', 'Qtd_IDs', 'Record_IDs', 'Etapas', 'Datas', 'Proprietarios']
duplicados_internos = duplicados_internos[duplicados_internos['Qtd_IDs'] > 1].sort_values('Qtd_IDs', ascending=False)

print(f"Total de combinações Nome+Unidade duplicadas DENTRO dos leads que sumiram: {len(duplicados_internos):,}")
print(f"Total de leads envolvidos nessas duplicatas: {duplicados_internos['Qtd_IDs'].sum():,}")

if len(duplicados_internos) > 0:
    print(f"\nTOP 15 CASOS DE MÚLTIPLOS IDs PARA MESMO NOME+UNIDADE:")
    print("-" * 100)
    
    for idx, row in duplicados_internos.head(15).iterrows():
        nome, unidade = row['Chave'].split('|||')
        print(f"\n{nome[:50]} - {unidade[:30]}")
        print(f"  • Qtd de Record IDs: {row['Qtd_IDs']}")
        print(f"  • IDs: {', '.join(map(str, row['Record_IDs'][:5]))}")
        print(f"  • Etapas: {', '.join(row['Etapas'][:3])}")
        print(f"  • Proprietários: {', '.join(row['Proprietarios'][:3])}")

# =======================================================================================
# CRITÉRIO 4: ANÁLISE TEMPORAL - IDs criados próximos podem ser duplicatas
# =======================================================================================
print("\n" + "="*100)
print("📋 CRITÉRIO 4: ANÁLISE TEMPORAL - LEADS CRIADOS EM DATAS PRÓXIMAS")
print("-" * 100)

# Para cada combinação Nome+Unidade, verificar se há leads criados com menos de 30 dias de diferença
duplicatas_temporais = []

for chave in duplicatas_nome_unidade:
    nome, unidade = chave.split('|||')
    
    leads_sumidos_chave = df_sumidos_valido[df_sumidos_valido['Chave_Nome_Unidade'] == chave].copy()
    leads_hoje_chave = df_hoje_valido[df_hoje_valido['Chave_Nome_Unidade'] == chave].copy()
    
    # Comparar datas de criação
    for _, lead_sumido in leads_sumidos_chave.iterrows():
        data_sumido = lead_sumido['Data de criação']
        
        for _, lead_hoje in leads_hoje_chave.iterrows():
            data_hoje = lead_hoje['Data de criação']
            
            diferenca_dias = abs((data_sumido - data_hoje).days)
            
            if diferenca_dias <= 30:  # Criados com menos de 30 dias de diferença
                duplicatas_temporais.append({
                    'Nome': nome[:40],
                    'Unidade': unidade[:25],
                    'ID Sumido': lead_sumido['Record ID'],
                    'ID Hoje': lead_hoje['Record ID'],
                    'Data Sumido': data_sumido.strftime('%Y-%m-%d'),
                    'Data Hoje': data_hoje.strftime('%Y-%m-%d'),
                    'Diferença (dias)': diferenca_dias,
                    'Etapa Sumida': lead_sumido['Etapa do negócio'],
                    'Etapa Hoje': lead_hoje['Etapa do negócio'],
                    'Prop Sumido': lead_sumido['Proprietário do negócio'],
                    'Prop Hoje': lead_hoje['Proprietário do negócio']
                })

df_temp = pd.DataFrame(duplicatas_temporais).sort_values('Diferença (dias)')

print(f"Leads com ALTA PROBABILIDADE de duplicata (criados com <30 dias de diferença): {len(df_temp):,}")

if len(df_temp) > 0:
    print(f"\nTOP 20 DUPLICATAS MAIS PROVÁVEIS (menor diferença de tempo):")
    print("-" * 100)
    for idx, row in df_temp.head(20).iterrows():
        print(f"{row['Nome'][:35]:<37} {row['Unidade'][:20]:<22} Δ={row['Diferença (dias)']:>3} dias")

# =======================================================================================
# RESUMO EXECUTIVO
# =======================================================================================
print("\n" + "="*100)
print("📊 RESUMO EXECUTIVO - ANÁLISE RIGOROSA DE DUPLICATAS")
print("="*100)

leads_duplicados_criterio1 = len(duplicatas_nome_unidade)
leads_duplicados_criterio2 = len(duplicatas_completas)
leads_duplicados_criterio3 = duplicados_internos['Qtd_IDs'].sum() if len(duplicados_internos) > 0 else 0
leads_duplicados_criterio4 = len(df_temp)

print(f"""
TOTAL DE LEADS QUE SUMIRAM: {len(df_sumidos):,}
Leads válidos (com nome e unidade): {len(df_sumidos_valido):,}

CRITÉRIO 1 - Nome + Unidade:
  • Duplicatas identificadas: {leads_duplicados_criterio1:,} ({leads_duplicados_criterio1/len(df_sumidos_valido)*100:.1f}%)

CRITÉRIO 2 - Nome + Unidade + Proprietário:
  • Duplicatas exatas: {leads_duplicados_criterio2:,} ({leads_duplicados_criterio2/len(df_sumidos_valido)*100:.1f}%)

CRITÉRIO 3 - Múltiplos IDs internos:
  • Leads duplicados dentro dos sumidos: {leads_duplicados_criterio3:,} ({leads_duplicados_criterio3/len(df_sumidos)*100:.1f}%)

CRITÉRIO 4 - Temporal (<30 dias):
  • Alta probabilidade de duplicata: {leads_duplicados_criterio4:,} ({leads_duplicados_criterio4/len(df_sumidos_valido)*100:.1f}%)

CONCLUSÃO:
""")

if leads_duplicados_criterio2 > len(df_sumidos) * 0.5:
    print("✅ MAIS DE 50% dos leads apresentam duplicatas comprovadas (Nome+Unidade+Proprietário).")
    print("   A limpeza de duplicatas é a causa PRINCIPAL da redução.")
elif leads_duplicados_criterio1 > len(df_sumidos) * 0.3:
    print(f"⚠️  {leads_duplicados_criterio1/len(df_sumidos_valido)*100:.1f}% dos leads são duplicatas por Nome+Unidade.")
    print("   A limpeza de duplicatas explica PARCIALMENTE a redução.")
else:
    print(f"❌ Apenas {leads_duplicados_criterio1/len(df_sumidos_valido)*100:.1f}% são duplicatas claras.")
    print("   A limpeza de duplicatas NÃO explica a maior parte da redução.")

# Salvar análise detalhada
print("\n" + "="*100)
print("💾 GERANDO ARQUIVO EXCEL COM ANÁLISE RIGOROSA...")
print("="*100)

with pd.ExcelWriter('ANALISE_DUPLICATAS_RIGOROSA.xlsx', engine='openpyxl') as writer:
    # Critério 1
    df_duplicatas_1.to_excel(writer, sheet_name='Critério 1 - Nome+Unidade', index=False)
    
    # Critério 3
    if len(duplicados_internos) > 0:
        duplicados_internos_export = duplicados_internos.copy()
        duplicados_internos_export['Record_IDs'] = duplicados_internos_export['Record_IDs'].apply(lambda x: ', '.join(map(str, x[:10])))
        duplicados_internos_export['Etapas'] = duplicados_internos_export['Etapas'].apply(lambda x: ', '.join(x))
        duplicados_internos_export['Proprietarios'] = duplicados_internos_export['Proprietarios'].apply(lambda x: ', '.join(x[:3]))
        duplicados_internos_export.to_excel(writer, sheet_name='Critério 3 - IDs Duplicados', index=False)
    
    # Critério 4
    if len(df_temp) > 0:
        df_temp.to_excel(writer, sheet_name='Critério 4 - Temporal', index=False)
    
    # Resumo
    resumo = pd.DataFrame({
        'Critério': [
            'Total de leads que sumiram',
            'Leads válidos (nome + unidade)',
            '',
            'CRITÉRIO 1: Nome + Unidade',
            'CRITÉRIO 2: Nome + Unidade + Proprietário',
            'CRITÉRIO 3: Múltiplos IDs internos',
            'CRITÉRIO 4: Temporal (<30 dias)',
        ],
        'Quantidade': [
            len(df_sumidos),
            len(df_sumidos_valido),
            '',
            leads_duplicados_criterio1,
            leads_duplicados_criterio2,
            leads_duplicados_criterio3,
            leads_duplicados_criterio4,
        ],
        'Percentual': [
            '100%',
            f"{len(df_sumidos_valido)/len(df_sumidos)*100:.1f}%",
            '',
            f"{leads_duplicados_criterio1/len(df_sumidos_valido)*100:.1f}%",
            f"{leads_duplicados_criterio2/len(df_sumidos_valido)*100:.1f}%",
            f"{leads_duplicados_criterio3/len(df_sumidos)*100:.1f}%",
            f"{leads_duplicados_criterio4/len(df_sumidos_valido)*100:.1f}%",
        ]
    })
    resumo.to_excel(writer, sheet_name='Resumo', index=False)

print("\n✅ ARQUIVO CRIADO: ANALISE_DUPLICATAS_RIGOROSA.xlsx")
print("\n📑 Abas criadas:")
print("   • Critério 1 - Nome+Unidade - Leads com mesmo nome e unidade")
print("   • Critério 3 - IDs Duplicados - Múltiplos IDs para mesma pessoa")
print("   • Critério 4 - Temporal - Leads criados em datas próximas")
print("   • Resumo - Consolidação de todos os critérios")
print("\n" + "="*100)
