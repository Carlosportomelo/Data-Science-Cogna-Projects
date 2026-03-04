import pandas as pd
import re
import unicodedata
from pathlib import Path

CORRECT_FILE = Path(r'C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_atual.csv')

def clean_cols(df):
    cols = []
    for c in df.columns:
        c_str = str(c).strip().lower()
        c_norm = unicodedata.normalize('NFKD', c_str)
        c_ascii = c_norm.encode('ascii', 'ignore').decode('utf-8')
        c_clean = re.sub(r'[^a-z0-9_ ]+', '', c_ascii)
        c_clean = re.sub(r'\s+', '_', c_clean)
        cols.append(c_clean)
    df.columns = cols
    return df

print(f"Verificando arquivo: {CORRECT_FILE}")
print(f"Existe? {CORRECT_FILE.exists()}\n")

if CORRECT_FILE.exists():
    df = pd.read_csv(CORRECT_FILE, encoding='utf-8', sep=None, engine='python')
    print(f"✓ Total de linhas: {len(df):,}")
    
    df = clean_cols(df)
    df = df.rename(columns={'data_de_criacao': 'Data'})
    df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
    
    # Janeiro 2026
    df_jan26 = df[(df['Data'] >= '2026-01-01') & (df['Data'] < '2026-02-01')]
    print(f"✓ Janeiro 2026: {len(df_jan26):,} leads")
    print(f"\nEsperado: 2,545 leads")
    print(f"Match: {'✓ SIM' if len(df_jan26) == 2545 else '✗ NÃO'}")
else:
    print("❌ ARQUIVO NÃO ENCONTRADO!")
