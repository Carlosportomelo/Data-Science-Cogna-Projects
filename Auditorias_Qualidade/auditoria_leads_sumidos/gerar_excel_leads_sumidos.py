import pandas as pd

print("Carregando dados...")
# Carregar os arquivos CSV já gerados
df_sumidos_completo = pd.read_csv('leads_sumidos_completo.csv')
df_matriculas = pd.read_csv('matriculas_sumidas.csv')
df_perdidos = pd.read_csv('negocios_perdidos_sumidos.csv')
df_novos = pd.read_csv('novos_negocios_sumidos.csv')

# Criar um resumo executivo
resumo_data = {
    'Categoria': [
        'TOTAL DE LEADS SUMIDOS',
        '',
        'Matrículas Concluídas',
        'Negócios Perdidos',
        'Novos Negócios',
        'Outros (Qualificação, Pausa, etc)',
        '',
        'DETALHAMENTO:',
        'Leads de 2021-2025 (antigas)',
        'Leads de Fevereiro/2026 (recentes)',
        '',
        'OBSERVAÇÃO:',
    ],
    'Quantidade': [
        len(df_sumidos_completo),
        '',
        len(df_matriculas),
        len(df_perdidos),
        len(df_novos),
        len(df_sumidos_completo) - len(df_matriculas) - len(df_perdidos) - len(df_novos),
        '',
        '',
        len(df_sumidos_completo[pd.to_datetime(df_sumidos_completo['Data de criação']) < '2026-01-01']),
        len(df_sumidos_completo[pd.to_datetime(df_sumidos_completo['Data de criação']) >= '2026-02-01']),
        '',
        '',
    ],
    'Percentual': [
        '100%',
        '',
        f"{len(df_matriculas)/len(df_sumidos_completo)*100:.1f}%",
        f"{len(df_perdidos)/len(df_sumidos_completo)*100:.1f}%",
        f"{len(df_novos)/len(df_sumidos_completo)*100:.1f}%",
        f"{(len(df_sumidos_completo) - len(df_matriculas) - len(df_perdidos) - len(df_novos))/len(df_sumidos_completo)*100:.1f}%",
        '',
        '',
        '',
        '',
        '',
        '',
    ],
    'Observação': [
        'Leads que estavam na base de sexta-feira (06/02) mas não estão na base de hoje (09/02)',
        '',
        'CRÍTICO: Matrículas concluídas não deveriam ser removidas do histórico',
        'Negócios perdidos podem ser arquivados conforme política da empresa',
        'Novos negócios removidos antes de qualquer ação',
        'Leads em processo que desapareceram',
        '',
        '',
        'Leads criadas entre 2021-2025 que foram removidas',
        'Leads muito recentes (criadas dias 3-5/fev) que já desapareceram',
        '',
        'Verificar urgentemente no HubSpot se houve exclusão em massa ou problema de sincronização',
    ]
}

df_resumo = pd.DataFrame(resumo_data)

# Análise por proprietário
analise_prop = df_sumidos_completo.groupby('Proprietário do negócio').agg({
    'Record ID': 'count',
    'Etapa do negócio': lambda x: x.value_counts().to_dict()
}).reset_index()
analise_prop.columns = ['Proprietário', 'Total de Leads Sumidos', 'Distribuição por Etapa']
analise_prop = analise_prop.sort_values('Total de Leads Sumidos', ascending=False)

# Análise por unidade
analise_unid = df_sumidos_completo.groupby('Unidade Desejada').agg({
    'Record ID': 'count',
    'Etapa do negócio': lambda x: x.value_counts().to_dict()
}).reset_index()
analise_unid.columns = ['Unidade', 'Total de Leads Sumidos', 'Distribuição por Etapa']
analise_unid = analise_unid.sort_values('Total de Leads Sumidos', ascending=False)

# Análise temporal
df_sumidos_completo['Data de criação'] = pd.to_datetime(df_sumidos_completo['Data de criação'])
df_sumidos_completo['Ano-Mês'] = df_sumidos_completo['Data de criação'].dt.to_period('M').astype(str)
analise_temporal = df_sumidos_completo.groupby('Ano-Mês').agg({
    'Record ID': 'count',
    'Etapa do negócio': lambda x: x.value_counts().to_dict()
}).reset_index()
analise_temporal.columns = ['Período (Ano-Mês)', 'Total de Leads Sumidos', 'Distribuição por Etapa']

print("Criando arquivo Excel...")
# Criar arquivo Excel com múltiplas abas
with pd.ExcelWriter('LEADS_SUMIDOS_ANALISE_COMPLETA.xlsx', engine='openpyxl') as writer:
    # Aba 1: Resumo Executivo
    df_resumo.to_excel(writer, sheet_name='RESUMO EXECUTIVO', index=False)
    
    # Aba 2: Todos os leads sumidos
    df_sumidos_completo.to_excel(writer, sheet_name='TODOS OS LEADS (1063)', index=False)
    
    # Aba 3: Matrículas concluídas
    df_matriculas.to_excel(writer, sheet_name='MATRÍCULAS SUMIDAS (485)', index=False)
    
    # Aba 4: Negócios perdidos
    df_perdidos.to_excel(writer, sheet_name='NEGÓCIOS PERDIDOS (348)', index=False)
    
    # Aba 5: Novos negócios
    df_novos.to_excel(writer, sheet_name='NOVOS NEGÓCIOS (147)', index=False)
    
    # Aba 6: Análise por proprietário
    analise_prop.to_excel(writer, sheet_name='Análise por Proprietário', index=False)
    
    # Aba 7: Análise por unidade
    analise_unid.to_excel(writer, sheet_name='Análise por Unidade', index=False)
    
    # Aba 8: Análise temporal
    analise_temporal.to_excel(writer, sheet_name='Análise Temporal', index=False)

print("\n" + "="*80)
print("✅ ARQUIVO EXCEL CRIADO COM SUCESSO!")
print("="*80)
print("\n📊 Arquivo: LEADS_SUMIDOS_ANALISE_COMPLETA.xlsx")
print("\n📑 Abas criadas:")
print("   1. RESUMO EXECUTIVO - Visão geral do problema")
print("   2. TODOS OS LEADS (1063) - Lista completa dos leads que sumiram")
print("   3. MATRÍCULAS SUMIDAS (485) - Matrículas concluídas que desapareceram")
print("   4. NEGÓCIOS PERDIDOS (348) - Negócios perdidos que foram removidos")
print("   5. NOVOS NEGÓCIOS (147) - Novos negócios que sumiram")
print("   6. Análise por Proprietário - Consolidação por responsável")
print("   7. Análise por Unidade - Consolidação por unidade")
print("   8. Análise Temporal - Distribuição por período de criação")
print("\n" + "="*80)
