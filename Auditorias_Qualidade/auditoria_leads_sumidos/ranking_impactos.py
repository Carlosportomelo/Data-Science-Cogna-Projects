import pandas as pd

print("Carregando dados dos leads sumidos...")
df_sumidos = pd.read_csv('leads_sumidos_completo.csv')

print("\n" + "="*100)
print("🏆 RANKING DE SECRETÁRIAS/PROPRIETÁRIOS MAIS IMPACTADOS")
print("="*100)

# Análise detalhada por proprietário
ranking_prop = df_sumidos.groupby('Proprietário do negócio').agg({
    'Record ID': 'count',
}).reset_index()
ranking_prop.columns = ['Proprietário', 'Total Leads Perdidos']

# Separar por etapa
matriculas_por_prop = df_sumidos[df_sumidos['Etapa do negócio'] == 'MATRÍCULA CONCLUÍDA (Red Balloon - Unidades de Rua)'].groupby('Proprietário do negócio')['Record ID'].count()
perdidos_por_prop = df_sumidos[df_sumidos['Etapa do negócio'] == 'NEGÓCIO PERDIDO (Red Balloon - Unidades de Rua)'].groupby('Proprietário do negócio')['Record ID'].count()
novos_por_prop = df_sumidos[df_sumidos['Etapa do negócio'] == 'NOVO NEGÓCIO (Red Balloon - Unidades de Rua)'].groupby('Proprietário do negócio')['Record ID'].count()

ranking_prop = ranking_prop.merge(matriculas_por_prop.rename('Matrículas'), left_on='Proprietário', right_index=True, how='left')
ranking_prop = ranking_prop.merge(perdidos_por_prop.rename('Perdidos'), left_on='Proprietário', right_index=True, how='left')
ranking_prop = ranking_prop.merge(novos_por_prop.rename('Novos'), left_on='Proprietário', right_index=True, how='left')

ranking_prop = ranking_prop.fillna(0).sort_values('Total Leads Perdidos', ascending=False)

# Converter para int
ranking_prop['Matrículas'] = ranking_prop['Matrículas'].astype(int)
ranking_prop['Perdidos'] = ranking_prop['Perdidos'].astype(int)
ranking_prop['Novos'] = ranking_prop['Novos'].astype(int)

print("\nTOP 30 SECRETÁRIAS/PROPRIETÁRIOS:")
print("-" * 100)
print(f"{'#':<4} {'Secretária/Proprietário':<50} {'Total':<8} {'Matrículas':<12} {'Perdidos':<10} {'Novos':<8}")
print("-" * 100)

for idx, row in ranking_prop.head(30).iterrows():
    pos = ranking_prop.index.get_loc(idx) + 1
    print(f"{pos:<4} {row['Proprietário'][:48]:<50} {row['Total Leads Perdidos']:<8} "
          f"{row['Matrículas']:<12} {row['Perdidos']:<10} "
          f"{row['Novos']:<8}")

print("\n" + "="*100)
print("🏢 RANKING DE UNIDADES MAIS IMPACTADAS")
print("="*100)

# Análise detalhada por unidade
ranking_unid = df_sumidos.groupby('Unidade Desejada').agg({
    'Record ID': 'count',
}).reset_index()
ranking_unid.columns = ['Unidade', 'Total Leads Perdidos']

# Separar por etapa
matriculas_por_unid = df_sumidos[df_sumidos['Etapa do negócio'] == 'MATRÍCULA CONCLUÍDA (Red Balloon - Unidades de Rua)'].groupby('Unidade Desejada')['Record ID'].count()
perdidos_por_unid = df_sumidos[df_sumidos['Etapa do negócio'] == 'NEGÓCIO PERDIDO (Red Balloon - Unidades de Rua)'].groupby('Unidade Desejada')['Record ID'].count()
novos_por_unid = df_sumidos[df_sumidos['Etapa do negócio'] == 'NOVO NEGÓCIO (Red Balloon - Unidades de Rua)'].groupby('Unidade Desejada')['Record ID'].count()

ranking_unid = ranking_unid.merge(matriculas_por_unid.rename('Matrículas'), left_on='Unidade', right_index=True, how='left')
ranking_unid = ranking_unid.merge(perdidos_por_unid.rename('Perdidos'), left_on='Unidade', right_index=True, how='left')
ranking_unid = ranking_unid.merge(novos_por_unid.rename('Novos'), left_on='Unidade', right_index=True, how='left')

ranking_unid = ranking_unid.fillna(0).sort_values('Total Leads Perdidos', ascending=False)

# Converter para int
ranking_unid['Matrículas'] = ranking_unid['Matrículas'].astype(int)
ranking_unid['Perdidos'] = ranking_unid['Perdidos'].astype(int)
ranking_unid['Novos'] = ranking_unid['Novos'].astype(int)

print("\nTOP 40 UNIDADES:")
print("-" * 100)
print(f"{'#':<4} {'Unidade':<40} {'Total':<8} {'Matrículas':<12} {'Perdidos':<10} {'Novos':<8}")
print("-" * 100)

for idx, row in ranking_unid.head(40).iterrows():
    pos = ranking_unid.index.get_loc(idx) + 1
    print(f"{pos:<4} {row['Unidade'][:38]:<40} {row['Total Leads Perdidos']:<8} "
          f"{row['Matrículas']:<12} {row['Perdidos']:<10} "
          f"{row['Novos']:<8}")

# Análise cruzada: Secretárias x Unidades mais críticas
print("\n" + "="*100)
print("🔍 ANÁLISE CRUZADA: SECRETÁRIAS COM MAIS MATRÍCULAS PERDIDAS POR UNIDADE")
print("="*100)

matriculas_perdidas = df_sumidos[df_sumidos['Etapa do negócio'] == 'MATRÍCULA CONCLUÍDA (Red Balloon - Unidades de Rua)']
cruzamento = matriculas_perdidas.groupby(['Unidade Desejada', 'Proprietário do negócio']).size().reset_index(name='Matrículas Perdidas')
cruzamento = cruzamento.sort_values('Matrículas Perdidas', ascending=False)

print("\nTOP 25 COMBINAÇÕES CRÍTICAS (Unidade + Secretária):")
print("-" * 100)
print(f"{'#':<4} {'Unidade':<30} {'Secretária/Proprietário':<45} {'Matrículas':<12}")
print("-" * 100)

for idx, row in cruzamento.head(25).iterrows():
    pos = cruzamento.index.get_loc(idx) + 1
    print(f"{pos:<4} {row['Unidade Desejada'][:28]:<30} {row['Proprietário do negócio'][:43]:<45} {row['Matrículas Perdidas']:<12}")

# Estatísticas gerais
print("\n" + "="*100)
print("📊 ESTATÍSTICAS GERAIS")
print("="*100)

total_secretarias = len(ranking_prop)
total_unidades = len(ranking_unid)
top10_secretarias = ranking_prop.head(10)['Total Leads Perdidos'].sum()
top10_unidades = ranking_unid.head(10)['Total Leads Perdidos'].sum()

print(f"""
Total de Secretárias/Proprietários Impactados: {total_secretarias}
Total de Unidades Impactadas: {total_unidades}

Top 10 Secretárias concentram: {top10_secretarias} leads ({top10_secretarias/len(df_sumidos)*100:.1f}% do total)
Top 10 Unidades concentram: {top10_unidades} leads ({top10_unidades/len(df_sumidos)*100:.1f}% do total)

Secretária MAIS impactada: {ranking_prop.iloc[0]['Proprietário']} ({ranking_prop.iloc[0]['Total Leads Perdidos']} leads)
Unidade MAIS impactada: {ranking_unid.iloc[0]['Unidade']} ({ranking_unid.iloc[0]['Total Leads Perdidos']} leads)
""")

# Criar arquivo Excel com os rankings
print("\n" + "="*100)
print("💾 GERANDO ARQUIVO EXCEL COM RANKINGS...")
print("="*100)

with pd.ExcelWriter('RANKING_SECRETARIAS_UNIDADES_IMPACTADAS.xlsx', engine='openpyxl') as writer:
    # Ranking de secretárias
    ranking_prop.to_excel(writer, sheet_name='Ranking Secretárias', index=False)
    
    # Ranking de unidades
    ranking_unid.to_excel(writer, sheet_name='Ranking Unidades', index=False)
    
    # Análise cruzada
    cruzamento.to_excel(writer, sheet_name='Análise Cruzada', index=False)
    
    # Top 10 detalhado
    top10_sec = ranking_prop.head(10)
    top10_sec.to_excel(writer, sheet_name='TOP 10 Secretárias', index=False)
    
    top10_uni = ranking_unid.head(10)
    top10_uni.to_excel(writer, sheet_name='TOP 10 Unidades', index=False)
    
    # Estatísticas
    stats_data = {
        'Métrica': [
            'Total de Leads Sumidos',
            'Total de Secretárias/Proprietários Impactados',
            'Total de Unidades Impactadas',
            '',
            'Secretária MAIS Impactada',
            'Leads Perdidos (Secretária #1)',
            '',
            'Unidade MAIS Impactada',
            'Leads Perdidos (Unidade #1)',
            '',
            'Top 10 Secretárias (% do total)',
            'Top 10 Unidades (% do total)',
        ],
        'Valor': [
            len(df_sumidos),
            total_secretarias,
            total_unidades,
            '',
            ranking_prop.iloc[0]['Proprietário'],
            ranking_prop.iloc[0]['Total Leads Perdidos'],
            '',
            ranking_unid.iloc[0]['Unidade'],
            ranking_unid.iloc[0]['Total Leads Perdidos'],
            '',
            f"{top10_secretarias/len(df_sumidos)*100:.1f}%",
            f"{top10_unidades/len(df_sumidos)*100:.1f}%",
        ]
    }
    df_stats = pd.DataFrame(stats_data)
    df_stats.to_excel(writer, sheet_name='Estatísticas', index=False)

print("\n✅ ARQUIVO CRIADO: RANKING_SECRETARIAS_UNIDADES_IMPACTADAS.xlsx")
print("\n📑 Abas criadas:")
print("   • Ranking Secretárias - Ranking completo de todos os proprietários")
print("   • Ranking Unidades - Ranking completo de todas as unidades")
print("   • Análise Cruzada - Combinações Unidade + Secretária mais críticas")
print("   • TOP 10 Secretárias - Top 10 detalhado")
print("   • TOP 10 Unidades - Top 10 detalhado")
print("   • Estatísticas - Resumo executivo")
print("\n" + "="*100)
