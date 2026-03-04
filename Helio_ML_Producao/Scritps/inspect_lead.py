import pandas as pd
import os, glob
from datetime import datetime
base = r"c:\Users\a483650\projeto_helio"
path = os.path.join(base, 'Outputs', 'Dados_Scored')
files = glob.glob(os.path.join(path, 'leads_scored_*.csv'))
if not files:
    print('Nenhum arquivo encontrado em', path)
    raise SystemExit
f = max(files, key=os.path.getctime)
print('Arquivo usado:', f)

df = pd.read_csv(f, sep=';', encoding='utf-8-sig', dtype=str)
rec_id = '48278926656'
row = df[df['Record ID'] == rec_id]
if row.empty:
    print('Record ID não encontrado:', rec_id)
else:
    raw = row['Data de criação'].iloc[0] if 'Data de criação' in row.columns else None
    print('Raw Data de criação:', repr(raw))
    print('Dias_Desde_Criacao (arquivo):', row['Dias_Desde_Criacao'].iloc[0] if 'Dias_Desde_Criacao' in row.columns else 'N/A')
    for dayfirst in (True, False):
        try:
            parsed = pd.to_datetime(raw, dayfirst=dayfirst, errors='coerce')
            print('dayfirst=', dayfirst, '-> parsed=', parsed)
            if pd.notna(parsed):
                diff = (pd.Timestamp(datetime.now().date()) - parsed.normalize()).days
                print('  diff days to today:', diff)
        except Exception as e:
            print('  parse error:', e)
