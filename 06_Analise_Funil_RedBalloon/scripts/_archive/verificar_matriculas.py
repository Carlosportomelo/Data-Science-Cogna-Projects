import pandas as pd
from pathlib import Path

excel_file = Path(r"c:\Users\a483650\Projetos\06_Analise_Funil_RedBalloon\outputs\funil_redballoon_20260211_174832.xlsx")

# Verifica aba comparativa
print("=" * 80)
print("ABA 7b_Comparativo_26vs25")
print("=" * 80)
df_comp = pd.read_excel(excel_file, sheet_name="7b_Comparativo_26vs25")
print(df_comp)

print("\n" + "=" * 80)
print("CHECANDO DADOS RAW DO ARQUIVO")
print("=" * 80)

# Verifica se há um CSV ou se preciso ler o CSV original
csv_file = Path(r"c:\Users\a483650\Projetos\06_Analise_Funil_RedBalloon\data\hubspot_leads_redballoon.csv")
if csv_file.exists():
    print(f"Arquivo CSV encontrado: {csv_file}")
    df_raw = pd.read_csv(csv_file)
    
    # Filtra Meta Ads no ciclo 25.1
    print("\n" + "=" * 80)
    print("META ADS - CICLO 25.1 (do CSV original)")
    print("=" * 80)
    
    if 'Fonte_Categoria' in df_raw.columns and 'lifecycle_stage' in df_raw.columns:
        meta_ads_251 = df_raw[(df_raw['Fonte_Categoria'] == 'Meta Ads') & (df_raw['Ciclo_Geracao'] == '25.1')]
        print(f"\nTotal de leads Meta Ads no 25.1: {len(meta_ads_251)}")
        
        if 'lifecycle_stage' in meta_ads_251.columns:
            print("\nDistribuição por status:")
            print(meta_ads_251['lifecycle_stage'].value_counts())
            
            matriculados = len(meta_ads_251[meta_ads_251['lifecycle_stage'] == 'customer'])
            print(f"\n✅ MATRÍCULAS Meta Ads no 25.1: {matriculados}")
else:
    print("CSV não encontrado. Verificando apenas Excel...")
