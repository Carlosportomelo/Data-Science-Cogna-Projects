import pandas as pd
import numpy as np
from datetime import datetime
import os

print("=" * 120)
print("COMPARAÇÃO DETALHADA: ANÁLISE 30/01 vs BASE DE BACKUP")
print("=" * 120)
print()

# Carregar análise do dia 30 - Resumo
analise_arquivo = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-01-30.xlsx'

print("Lendo abas da análise consolidada do dia 30...")
print()

# Aba Resumo
df_resumo = pd.read_excel(analise_arquivo, sheet_name='Resumo e Insights')
print("RESUMO E INSIGHTS:")
print(df_resumo.to_string(index=False))
print()

# Aba Matrículas por Canais
df_matriculas_canais = pd.read_excel(analise_arquivo, sheet_name='Matrículas (Canais)')
print("\nMATRÍCULAS POR CANAIS:")
print(df_matriculas_canais.to_string(index=False))
print()

# Aba Matrículas por Unidades
df_matriculas_unidades = pd.read_excel(analise_arquivo, sheet_name='Matrículas (Unidades)')
print("\nMATRÍCULAS POR UNIDADES (Top 15):")
print(df_matriculas_unidades.head(15).to_string(index=False))
print()

# Aba Leads por Canais
df_leads_canais = pd.read_excel(analise_arquivo, sheet_name='Leads (Canais)')
print("\nLEADS POR CANAIS:")
print(df_leads_canais.to_string(index=False))
print()

# Aba Leads por Unidades
df_leads_unidades = pd.read_excel(analise_arquivo, sheet_name='Leads (Unidades)')
print("\nLEADS POR UNIDADES (Top 15):")
print(df_leads_unidades.head(15).to_string(index=False))
print()

# Agora vamos carregar a base de backup e refazer os mesmos cálculos
print("\n" + "=" * 120)
print("RECRIANDO ANÁLISE USANDO BASE DE BACKUP (até 30/01/2026)")
print("=" * 120)
print()

base_backup = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads.csv'
df_backup = pd.read_csv(base_backup)
df_backup['Data de criação'] = pd.to_datetime(df_backup['Data de criação'], format='%Y-%m-%d %H:%M')
df_backup['Data de fechamento'] = pd.to_datetime(df_backup['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')

# Filtrar até 30/01/2026
data_corte = pd.Timestamp('2026-01-30 23:59:59')
df_backup_corte = df_backup[df_backup['Data de criação'] <= data_corte].copy()

print(f"Base de backup até 30/01/2026: {len(df_backup_corte)} registros")

# Definir ciclos (mesmo padrão da análise - provavelmente ciclos de 6 meses)
ciclos = {}
for yy in ['23', '24', '25', '26']:
    label = f"{yy}.1"
    start_year = 2000 + int(yy) - 1
    start = pd.Timestamp(f"{start_year}-10-01")
    end = pd.Timestamp(f"{start_year + 1}-02-20") + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    ciclos[label] = (start, end)

print("\nDados por ciclo conforme padrão da análise:")

resumo_ciclos = []
for ciclo, (data_inicio, data_fim) in ciclos.items():
    df_ciclo = df_backup_corte[(df_backup_corte['Data de criação'] >= data_inicio) & 
                                (df_backup_corte['Data de criação'] <= data_fim)].copy()
    
    df_matriculas_ciclo = df_ciclo[df_ciclo['Etapa do negócio'].astype(str).str.contains("MATRÍCULA CONCLUÍDA", case=False, na=False)]
    
    qtd_leads = len(df_ciclo)
    qtd_matriculas = len(df_matriculas_ciclo)
    
    resumo_ciclos.append({
        'Ciclo': ciclo,
        'Leads': qtd_leads,
        'Matrículas': qtd_matriculas,
        'Data Início': data_inicio.strftime('%Y-%m-%d'),
        'Data Fim': data_fim.strftime('%Y-%m-%d')
    })
    
    print(f"\n{ciclo} ({data_inicio.strftime('%Y-%m-%d')} a {data_fim.strftime('%Y-%m-%d')})")
    print(f"  Leads criados: {qtd_leads}")
    print(f"  Matrículas: {qtd_matriculas}")

df_resumo_recriado = pd.DataFrame(resumo_ciclos)
print("\n" + "-" * 120)
print("TABELA COMPARATIVA DE CICLOS:")
print("-" * 120)
print(df_resumo_recriado.to_string(index=False))

# Comparação com o arquivo do dia 30
print("\n\n" + "=" * 120)
print("COMPARAÇÃO COM ANÁLISE DO DIA 30:")
print("=" * 120)

print("\nResumo do dia 30 (arquivo Excel):")
print(df_resumo.to_string(index=False))

print("\nResumo recriado com base de backup:")
print(df_resumo_recriado[['Ciclo', 'Leads', 'Matrículas']].to_string(index=False))

# Exportar comparação
output_file = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\relatorio_auditoria_dados\comparacao_analise_30_backup.txt'
os.makedirs(os.path.dirname(output_file), exist_ok=True)

with open(output_file, 'w', encoding='utf-8') as f:
    f.write("COMPARAÇÃO DETALHADA: ANÁLISE 30/01 vs BASE DE BACKUP\n")
    f.write("=" * 120 + "\n")
    f.write(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n")
    
    f.write("ANÁLISE DO DIA 30/01 (arquivo Excel):\n")
    f.write(df_resumo.to_string(index=False))
    
    f.write("\n\nANÁLISE RECRIADA COM BASE DE BACKUP:\n")
    f.write(df_resumo_recriado[['Ciclo', 'Leads', 'Matrículas']].to_string(index=False))

print(f"\n✓ Comparação exportada para: {output_file}")
