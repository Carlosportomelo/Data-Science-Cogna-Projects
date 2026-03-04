´import pandas as pd
from pathlib import Path

# Define unidades próprias - São apenas 8 (Mooca é FRANQUEADA)
unidades_proprias_norm = {'vila leopoldina', 'morumbi', 'pacaembu', 'itaim bibi', 'pinheiros', 'jardins', 'perdizes', 'santana'}

# Carrega o arquivo principal
data_path = Path(r'c:\Users\a483650\Projetos\analise_eficiencia_canal\data')
hubspot_main = data_path / 'hubspot_leads.csv'

print("Carregando dados...")
try:
    df = pd.read_csv(hubspot_main)
    if len(df.columns) <= 1:
        df = pd.read_csv(hubspot_main, sep=';')
except:
    df = pd.read_csv(hubspot_main, sep=';')

df.columns = df.columns.str.strip()

print(f"Colunas disponíveis: {df.columns.tolist()}")

# Normalização de colunas
col_map = {
    'Data de criação': ['Data de criação', 'Create Date', 'Create date', 'Created Date'],
    'Unidade': ['Unidade', 'Campus', 'Unidade de interesse', 'Business Unit', 'BU', 'unidade', 'campus', 
                'Unidade desejada', 'unidade desejada', 'Unidade Desejada', 'Campus de interesse', 'Unidade de Negócio', 'Unidade de negócio']
}

actual_cols = {}
for key, alternatives in col_map.items():
    for alt in alternatives:
        if alt in df.columns:
            actual_cols[key] = alt
            break

df = df.rename(columns={v: k for k, v in actual_cols.items()})

# Converte data
if 'Data de criação' in df.columns:
    df['Data de criação'] = pd.to_datetime(df['Data de criação'], dayfirst=False, errors='coerce')

# Normaliza unidade
if 'Unidade' in df.columns:
    df['Unidade'] = df['Unidade'].astype(str).str.strip().str.title()
    df['Unidade'] = df['Unidade'].replace(['Nan', 'None', 'NaT'], pd.NA)

# Filtra ciclo 26.1
start_26 = pd.Timestamp('2025-10-01')
end_26 = pd.Timestamp('2026-02-06') + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

mask_26 = (df['Data de criação'] >= start_26) & (df['Data de criação'] <= end_26)
df_26 = df[mask_26].copy()

print(f"\nTotal de registros no ciclo 26.1: {len(df_26)}")

# Conta unidades únicas
if 'Unidade' in df_26.columns:
    unidades_unicas = df_26['Unidade'].dropna().unique()
    unidades_unicas = [u for u in unidades_unicas if u and str(u).strip() and str(u).lower() not in ['nan', 'none', 'sem unidade']]
    
    # Classifica
    proprias = [u for u in unidades_unicas if u.lower().strip() in unidades_proprias_norm]
    franqueadas = [u for u in unidades_unicas if u.lower().strip() not in unidades_proprias_norm]
    
    print(f"\n{'='*80}")
    print("CICLO 26.1 - ANÁLISE DE UNIDADES")
    print(f"{'='*80}")
    
    print(f"\nUnidades Próprias: {len(proprias)}")
    for u in sorted(proprias):
        count = df_26[df_26['Unidade'] == u].shape[0]
        print(f"  - {u}: {count} registros")
    
    print(f"\nUnidades Franqueadas: {len(franqueadas)}")
    for u in sorted(franqueadas):
        count = df_26[df_26['Unidade'] == u].shape[0]
        print(f"  - {u}: {count} registros")
    
    print(f"\n{'='*80}")
    print(f"TOTAL UNIDADES NO CICLO 26.1:")
    print(f"  Próprias: {len(proprias)}")
    print(f"  Franqueadas: {len(franqueadas)}")
    print(f"  Total Geral: {len(proprias) + len(franqueadas)}")
    print(f"{'='*80}")
else:
    print("Coluna 'Unidade' não encontrada!")
