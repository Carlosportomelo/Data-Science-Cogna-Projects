import pandas as pd
import numpy as np
from datetime import datetime
import os

print("=" * 130)
print("VALIDAÇÃO: ANÁLISE DO DIA 30 GERADA COM BASE DE BACKUP")
print("Comparar números da análise 30/01 com base de backup ATUAL")
print("=" * 130)
print()

# Carregar ambos os dados
analise_arquivo = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-01-30.xlsx'
base_backup = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads.csv'

# Ler análise do dia 30
df_analise_30 = pd.read_excel(analise_arquivo, sheet_name='Resumo e Insights')

print("ANÁLISE DO DIA 30/01/2026:")
print(df_analise_30.to_string(index=False))
print()

# Carregar base de backup
print("Carregando base de backup...")
df_backup = pd.read_csv(base_backup)
df_backup['Data de criação'] = pd.to_datetime(df_backup['Data de criação'], format='%Y-%m-%d %H:%M')
df_backup['Data de fechamento'] = pd.to_datetime(df_backup['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')

print(f"✓ {len(df_backup)} registros carregados\n")

# (recalculo e comparação mantidos no arquivo original)

output_file = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\relatorio_auditoria_dados\analise_30_validacao_backup.txt'
os.makedirs(os.path.dirname(output_file), exist_ok=True)

with open(output_file, 'w', encoding='utf-8') as f:
    f.write("VALIDAÇÃO: ANÁLISE 30/01 COM BASE DE BACKUP\n")
    f.write("=" * 130 + "\n")
    f.write(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n")
    f.write("RESULTADOS:\n")
    f.write("-" * 130 + "\n\n")

print(f"✓ Relatório exportado para: {output_file}")
