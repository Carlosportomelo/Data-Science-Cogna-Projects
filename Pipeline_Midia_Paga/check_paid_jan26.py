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

def clean_text(text):
    try:
        if pd.isna(text):
            return ''
        text_str = str(text).lower().strip()
        text_norm = unicodedata.normalize('NFKD', text_str)
        text_ascii = text_norm.encode('ascii', 'ignore').decode('utf-8')
        return text_ascii
    except:
        return ''

df = pd.read_csv(CORRECT_FILE, encoding='utf-8', sep=None, engine='python')
print(f"Total de linhas: {len(df):,}\n")

df = clean_cols(df)
df = df.rename(columns={
    'data_de_criacao': 'Data',
    'fonte_original_do_trafego': 'Fonte_Original_do_Trafego'
})
df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
df['Fonte_clean'] = df['Fonte_Original_do_Trafego'].apply(clean_text)

# Classificar canal
KEYWORDS_SOCIAL = ['social pago', 'paid social', 'facebook', 'instagram', 'linkedin']
KEYWORDS_SEARCH = ['pesquisa paga', 'paid search', 'cpc', 'google ads']

def classify_channel(text):
    t = str(text).lower()
    if any(k in t for k in KEYWORDS_SOCIAL): return 'Social Pago'
    if any(k in t for k in KEYWORDS_SEARCH): return 'Pesquisa Paga'
    return 'Orgânico/Outros'

df['Origem_Principal'] = df['Fonte_clean'].apply(classify_channel)

# Janeiro 2026
df_jan26 = df[(df['Data'] >= '2026-01-01') & (df['Data'] < '2026-02-01')].copy()

print("JANEIRO 2026:")
print(f"  Total: {len(df_jan26):,}")
print(f"\nPor Canal:")
for canal in ['Social Pago', 'Pesquisa Paga', 'Orgânico/Outros']:
    count = len(df_jan26[df_jan26['Origem_Principal'] == canal])
    print(f"  {canal}: {count:,}")

# Total mídia paga
pago = df_jan26[df_jan26['Origem_Principal'].isin(['Social Pago', 'Pesquisa Paga'])]
print(f"\n  MÍDIA PAGA (Social + Pesquisa): {len(pago):,}")
print(f"\nEsperado pelo usuário: 2,545")
