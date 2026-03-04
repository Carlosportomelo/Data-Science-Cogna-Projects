import pandas as pd
from pathlib import Path

# Verificar qual arquivo está sendo usado
base_dir = Path('.')
data_dir = base_dir / 'data'

hubspot_file = data_dir / "hubspot_dataset.csv"
if not hubspot_file.exists():
    if (data_dir / "hubspot_leads.csv").exists():
        hubspot_file = data_dir / "hubspot_leads.csv"

print(f"Arquivo sendo usado: {hubspot_file}")
print(f"Existe? {hubspot_file.exists()}")

if hubspot_file.exists():
    df = pd.read_csv(hubspot_file, encoding='utf-8-sig', low_memory=False)
    print(f"\nTotal de linhas: {len(df):,}")
    
    # Verificar janeiro 2026 por data de criação
    df.columns = df.columns.str.lower().str.strip()
    data_col = None
    for col in df.columns:
        if 'data' in col and 'criacao' in col or 'data_de_criacao' in col:
            data_col = col
            break
            
    if data_col:
        print(f"Coluna de data encontrada: {data_col}")
        df['data'] = pd.to_datetime(df[data_col], errors='coerce')
        df['ano'] = df['data'].dt.year
        df['mes'] = df['data'].dt.month
        
        jan_2026 = df[(df['ano'] == 2026) & (df['mes'] == 1)]
        print(f"\nJaneiro 2026 (por data de criação): {len(jan_2026):,} leads")
    else:
        print("Coluna de data não encontrada")
        print(f"Colunas disponíveis: {list(df.columns)[:10]}")
