import pandas as pd
import numpy as np
from datetime import datetime
import os

output_dir = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\relatorio_auditoria_dados'
os.makedirs(output_dir, exist_ok=True)

# Dados da análise do dia 30
dados_analise_30 = {
    'Ciclo': ['23.1', '24.1', '25.1', '26.1'],
    'Leads_Analise30': [6290, 8296, 9688, 10421],
    'Matriculas_Analise30': [1342, 2138, 2800, 2519]
}

# Dados recreados da base de backup
dados_backup = {
    'Ciclo': ['23.1', '24.1', '25.1', '26.1'],
    'Leads_Backup': [9073, 14691, 17987, 15127],
    'Matriculas_Backup': [2925, 4887, 2226, 4250]
}

df_comparacao = pd.DataFrame(dados_analise_30)
df_backup_df = pd.DataFrame(dados_backup)
df_final = df_comparacao.merge(df_backup_df, on='Ciclo')

# (rest of report generation)
relatorio_file = os.path.join(output_dir, 'AUDITORIA_FINAL_ANALISE_30_JAN.txt')

with open(relatorio_file, 'w', encoding='utf-8') as f:
    f.write("AUDITORIA FINAL - resumo\n")

print(f"✓ Relatório de auditoria final gerado: {relatorio_file}")
