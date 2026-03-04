import pandas as pd
from pathlib import Path

# Define unidades próprias - São apenas 8 (Mooca é FRANQUEADA)
unidades_proprias_norm = {'vila leopoldina', 'morumbi', 'pacaembu', 'itaim bibi', 'pinheiros', 'jardins', 'perdizes', 'santana'}

# Localiza o arquivo principal a partir da raiz do projeto (duas pastas acima deste arquivo)
repo_root = Path(__file__).resolve().parents[2]
candidates = [
    Path(r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_negocios_perdidos_ATUAL.csv"),
    repo_root / 'data' / 'hubspot_leads.csv',
    Path.cwd() / 'data' / 'hubspot_leads.csv',
]

hubspot_main = None
for p in candidates:
    if p.exists():
        hubspot_main = p
        break

if hubspot_main is None:
    raise FileNotFoundError(f"Arquivo 'hubspot_leads.csv' não encontrado. Procurado em: {', '.join(str(p) for p in candidates)}")

print(f"Carregando dados de: {hubspot_main}")
# Tenta carregar com separador padrão e, se necessário, com ponto-e-vírgula
try:
    df = pd.read_csv(hubspot_main)
    if len(df.columns) <= 1:
        df = pd.read_csv(hubspot_main, sep=';')
except Exception:
    df = pd.read_csv(hubspot_main, sep=';')

df.columns = df.columns.str.strip()

# Normalização de colunas
col_map = {
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

# Normaliza unidade
if 'Unidade' in df.columns:
    df['Unidade'] = df['Unidade'].astype(str).str.strip().str.title()
    df['Unidade'] = df['Unidade'].replace(['Nan', 'None', 'NaT'], pd.NA)
    
    # Pega todas as unidades únicas
    unidades_unicas = df['Unidade'].dropna().unique()
    unidades_unicas = [u for u in unidades_unicas if u and str(u).strip() and str(u).lower() not in ['nan', 'none', 'sem unidade']]
    
    # Classifica
    proprias = sorted([u for u in unidades_unicas if u.lower().strip() in unidades_proprias_norm])
    franqueadas = sorted([u for u in unidades_unicas if u.lower().strip() not in unidades_proprias_norm])
    
    print(f"\n{'='*80}")
    print("CLASSIFICAÇÃO DE TODAS AS UNIDADES")
    print(f"{'='*80}")
    
    print(f"\n🏢 UNIDADES PRÓPRIAS: {len(proprias)}")
    print("-" * 80)
    for i, u in enumerate(proprias, 1):
        count = df[df['Unidade'] == u].shape[0]
        print(f"  {i}. {u} ({count:,} registros)")
    
    print(f"\n🏪 UNIDADES FRANQUEADAS: {len(franqueadas)}")
    print("-" * 80)
    for i, u in enumerate(franqueadas, 1):
        count = df[df['Unidade'] == u].shape[0]
        print(f"  {i}. {u} ({count:,} registros)")
    
    print(f"\n{'='*80}")
    print("RESUMO GERAL:")
    print(f"  Unidades Próprias: {len(proprias)}")
    print(f"  Unidades Franqueadas: {len(franqueadas)}")
    print(f"  TOTAL DE UNIDADES: {len(proprias) + len(franqueadas)}")
    print(f"{'='*80}")
    
    # Mostra quais são as próprias
    print(f"\n💡 AS 8 UNIDADES PRÓPRIAS SÃO (Mooca é FRANQUEADA):")
    for u in proprias:
        print(f"   ✓ {u}")
else:
    print("Coluna 'Unidade' não encontrada!")
