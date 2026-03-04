import pandas as pd

excel_30 = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-01-30.xlsx'
excel_06 = r'c:\Users\a483650\Projetos\analise_eficiencia_canal\outputs\analise_consolidada_leads_matriculas\analise_consolidada_leads_matriculas_2026-02-06.xlsx'

# Carregar Leads (Canais) de ambos
df_leads_30 = pd.read_excel(excel_30, sheet_name='Leads (Canais)')
df_leads_06 = pd.read_excel(excel_06, sheet_name='Leads (Canais)')

print("=" * 130)
print("COMPARACAO: Excel 30/01 vs Excel 06/02 - Sheet 'Leads (Canais)'")
print("=" * 130 + "\n")

print("CICLO | 30/01 | 06/02 | Diferenca | Status")
print("─" * 130)

for ciclo in ['23.1', '24.1', '25.1', '26.1']:
    val_30 = df_leads_30[ciclo].sum()
    val_06 = df_leads_06[ciclo].sum()
    diff = val_06 - val_30
    pct = (diff / val_30) * 100 if val_30 != 0 else 0
    
    status = "OK" if diff >= 0 else "REDUCCAO!"
    print(f"{ciclo:5} | {val_30:5,.0f} | {val_06:5,.0f} | {diff:+7,.0f} ({pct:+5.1f}%) | {status}")

print("\n" + "=" * 130)
print("QUAIS CANAIS MUDARAM ENTRE 30/01 E 06/02?")
print("=" * 130 + "\n")

# Detalhe por canal
for ciclo in ['25.1']:  # Focar no ciclo problematico
    print(f"\nCICLO {ciclo} - Mudancas nos Canais:")
    print("─" * 130)
    print(f"{'Canal':<35} | {'30/01':>8} | {'06/02':>8} | {'Diferenca':>10}")
    print("─" * 130)
    
    for idx, row in df_leads_30.iterrows():
        canal = row['Fonte original do tráfego']
        val_30 = row[ciclo]
        val_06 = df_leads_06[df_leads_06['Fonte original do tráfego'] == canal][ciclo].values
        
        if len(val_06) > 0:
            val_06 = val_06[0]
            diff = val_06 - val_30
            
            if diff != 0:
                print(f"{canal:<35} | {val_30:>8,.0f} | {val_06:>8,.0f} | {diff:>+10,.0f}")
    
    total_30 = df_leads_30[ciclo].sum()
    total_06 = df_leads_06[ciclo].sum()
    print("─" * 130)
    print(f"{'TOTAL':<35} | {total_30:>8,.0f} | {total_06:>8,.0f} | {total_06 - total_30:>+10,.0f}")
