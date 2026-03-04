import csv
import os

CAMINHO_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PATH = os.path.join(CAMINHO_BASE, 'Data', 'hubspot_leads.csv')

bad_lines = []
record_count = 0
with open(PATH, 'r', encoding='utf-8', errors='replace', newline='') as f:
    reader = csv.reader(f)
    try:
        header = next(reader)
    except StopIteration:
        print('Arquivo vazio')
        raise SystemExit(1)
    ncols = len(header)
    for row in reader:
        record_count += 1
        if len(row) != ncols:
            # We keep the physical file line number unknown here since csv.reader
            # already coalesces multiline records; just store the differing row
            bad_lines.append((record_count+1, len(row), row[:5]))

print('Colunas esperadas (pelo header):', ncols)
print('Registros lógicos detectados pelo csv.reader (excl header):', record_count)
print('Total de linhas com número de colunas diferente do header:', len(bad_lines))
if bad_lines:
    print('\nExemplos (registro_index, ncols_detectadas, primeiros campos):')
    for t in bad_lines[:10]:
        print(t)
