import pandas as pd

df = pd.read_csv('output/days_invest_no_leads.csv')
df['date'] = pd.to_datetime(df['date'])
df_periodo = df[(df['date'] >= '2025-10-01') & (df['date'] <= '2026-03-02')].sort_values('date')

print(f'Total: {len(df_periodo)} dias com investimento sem leads (out/2025 a 02/03/2026)\n')

for _, r in df_periodo.iterrows():
    print(f"{r['date'].strftime('%d/%m/%Y')} - R$ {r['total_investimento']:,.2f}")
