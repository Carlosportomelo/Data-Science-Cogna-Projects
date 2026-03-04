import pandas as pd

df = pd.read_csv(r'C:\Users\a483650\Projetos\06_Analise_Funil_RedBalloon\data\hubspot_leads_redballoon.csv')

print("COLUNAS DISPONÍVEIS:")
print("=" * 80)
for i, col in enumerate(df.columns, 1):
    print(f"{i:2d}. {col}")

print("\n" + "=" * 80)
print("Colunas com 'atividade', 'última', 'last':")
print("=" * 80)
atividade_cols = [col for col in df.columns if 'atividade' in col.lower() or 'última' in col.lower() or 'ultimo' in col.lower() or 'last' in col.lower()]
for col in atividade_cols:
    print(f"✓ {col}")

print("\n" + "=" * 80)
print("Exemplo de dados das colunas relacionadas a atividade:")
print("=" * 80)
if atividade_cols:
    print(df[atividade_cols].head(3))
