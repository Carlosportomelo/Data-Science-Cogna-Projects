import pandas as pd
from datetime import datetime
import os

output_file = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\relatorio_auditoria_dados\RELATORIO_FINAL_MANIPULACAO_CORRIGIDO.txt'
os.makedirs(os.path.dirname(output_file), exist_ok=True)

with open(output_file, 'w', encoding='utf-8') as f:
    f.write("╔" + "═" * 158 + "╗\n")
    f.write("║" + " " * 158 + "║\n")
    f.write("║" + "RELATÓRIO FINAL - AUDITORIA FORENSE DE MANIPULAÇÃO HUBSPOT".center(158) + "║\n")
    f.write("║" + "Base Backup 30/01/2026 vs Base Atual 06/02/2026".center(158) + "║\n")
    f.write("║" + " " * 158 + "║\n")
    f.write("╚" + "═" * 158 + "╝\n\n")
    
    f.write(f"Data do Relatório: {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}\n")
    f.write(f"Período Analisado: Ciclos Out-Jan (23.1, 24.1, 25.1, 26.1)\n\n")
    
    # Simplified conclusions
    f.write("STATUS: 🔴 MANIPULAÇÃO CONFIRMADA\n\n")
    f.write("Foram encontradas evidências CRÍTICAS de manipulação...\n")
    f.write("FIM DO RELATÓRIO\n")

print("✓ Relatório final consolidado gerado!")
