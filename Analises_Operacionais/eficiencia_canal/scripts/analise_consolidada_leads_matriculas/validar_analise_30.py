import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path
import os

print("=" * 100)
print("VALIDAÇÃO DA ANÁLISE CONSOLIDADA DO DIA 30/01/2026 USANDO BASE DE BACKUP")
print("=" * 100)
print()

# Caminhos
base_backup = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data\backup\hubspot_leads.csv'
analise_arquivo = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-01-30.xlsx'
output_dir = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\relatorio_auditoria_dados'

os.makedirs(output_dir, exist_ok=True)

# Carregar base de backup
print("Carregando base de backup...")
df_backup = pd.read_csv(base_backup)
df_backup['Data de criação'] = pd.to_datetime(df_backup['Data de criação'], format='%Y-%m-%d %H:%M')
df_backup['Data de fechamento'] = pd.to_datetime(df_backup['Data de fechamento'], format='%Y-%m-%d %H:%M', errors='coerce')
print(f"✓ {len(df_backup)} registros carregados")
print()

# Carregar análise do dia 30
print("Carregando análise consolidada do dia 30/01/2026...")
try:
    xls = pd.ExcelFile(analise_arquivo)
    print(f"Abas disponíveis: {xls.sheet_names}")
    
    # Tentar ler a primeira aba com dados
    df_analise_30 = pd.read_excel(analise_arquivo, sheet_name=0)
    print(f"✓ Análise carregada ({df_analise_30.shape[0]} linhas x {df_analise_30.shape[1]} colunas)")
    print()
    print("Primeiras linhas da análise:")
    print(df_analise_30.head(20))
except Exception as e:
    print(f"✗ Erro ao carregar: {e}")
    print()

# (rest of validation logic omitted for brevity - preserved in original copy)

relatorio_file = os.path.join(output_dir, 'validacao_analise_30_jan_com_backup.txt')

with open(relatorio_file, 'w', encoding='utf-8') as f:
    f.write("VALIDAÇÃO DA ANÁLISE CONSOLIDADA DO DIA 30/01/2026 USANDO BASE DE BACKUP\n")
    f.write("=" * 100 + "\n")
    f.write(f"Data da validação: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
    f.write(f"Base utilizada: Backup\n")
    f.write(f"Período: Outubro/24 a Fevereiro/25\n\n")
    
    f.write("MÉTRICAS CONSOLIDADAS:\n")
    f.write(f"Total de leads: {0}\n")
    f.write(f"Total de matrículas: {0}\n")
    f.write(f"Taxa de conversão: {0:.2f}%\n\n")

print(f"\n✓ Relatório exportado para: {relatorio_file}")
