import pandas as pd

output_file = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\relatorio_auditoria_dados\AUDITORIA_FINAL_MANIPULACAO_CONFIRMADA.txt'

print("Gerando relatorio final...")

with open(output_file, 'w', encoding='utf-8') as f:
    f.write("RELATORIO FINAL DE AUDITORIA - MANIPULACAO CONFIRMADA\n")

print(f"Relatório salvo em: {output_file}")
