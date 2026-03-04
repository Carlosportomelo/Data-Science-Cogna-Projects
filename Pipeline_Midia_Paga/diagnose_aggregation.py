import pandas as pd
from pathlib import Path
import re
import unicodedata

DATA_DIR = Path('data')
HUBSPOT_FILE = DATA_DIR / "hubspot_dataset.csv"
if not HUBSPOT_FILE.exists():
    if (DATA_DIR / "hubspot_leads.csv").exists():
        HUBSPOT_FILE = DATA_DIR / "hubspot_leads.csv"

def clean_cols(df):
    """Limpa e normaliza os nomes das colunas de um DataFrame."""
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

print(f"Carregando arquivo: {HUBSPOT_FILE}")
df = pd.read_csv(HUBSPOT_FILE, encoding='utf-8', sep=None, engine='python')
print(f"✓ Total de linhas no arquivo bruto: {len(df):,}")

df = clean_cols(df)
print(f"✓ Após clean_cols: {len(df):,}")

# Renomear data_de_criacao para Data
df = df.rename(columns={'data_de_criacao': 'Data'})
df['Data'] = pd.to_datetime(df['Data'], errors='coerce')

# Filtro Janeiro 2026
df_jan26 = df[(df['Data'] >= '2026-01-01') & (df['Data'] < '2026-02-01')]
print(f"\n✓ Leads com Data de Criação em JANEIRO 2026 (ANTES deduplicação): {len(df_jan26):,}")

# Simular deduplicação por Lead_Key (record_id)
if 'record_id' in df.columns:
    df['Lead_Key'] = df['record_id'].astype(str)
    n_antes = len(df)
    df = df.drop_duplicates(subset=['Lead_Key'], keep='last')
    n_apos = len(df)
    print(f"✓ Após deduplicação por Lead_Key: {n_antes:,} → {n_apos:,} (removidos: {n_antes - n_apos:,})")
    
    df_jan26_dedup = df[(df['Data'] >= '2026-01-01') & (df['Data'] < '2026-02-01')]
    print(f"✓ Janeiro 2026 após deduplicação: {len(df_jan26_dedup):,}")

# Verificar quantos leads têm Data válida (não NaT)
df_with_date = df[df['Data'].notna()]
print(f"\n✓ Leads com Data de Criação válida (não NaT): {len(df_with_date):,} de {len(df):,}")
print(f"✓ Leads com Data NULA (NaT): {len(df) - len(df_with_date):,}")

# Janeiro 2026 final
df_jan26_final = df_with_date[(df_with_date['Data'] >= '2026-01-01') & (df_with_date['Data'] < '2026-02-01')]
print(f"\n✓ JANEIRO 2026 FINAL (com data válida): {len(df_jan26_final):,}")
print(f"\nEsperado pelo usuário: 2,545 leads")
print(f"Diferença: {2545 - len(df_jan26_final):,}")
