import os
import pandas as pd
from datetime import datetime

# Caminho para o dataset do Meta (relativo ao arquivo do script)
META_CSV = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', 'data', 'meta_dataset.csv'))

# Função para responder perguntas sobre investimento em Meta

def investimento_janeiro_2026_meta():
    try:
        df = pd.read_csv(META_CSV)

        # Detecta coluna de data e de valor automaticamente
        cols = [c for c in df.columns]
        date_candidates = ['data', 'Data', 'Dia', 'dia']
        value_candidates = ['investimento', 'Investimento', 'Valor usado (BRL)', 'valor usado (brl)']

        date_col = None
        value_col = None
        for c in cols:
            if date_col is None and c.strip().lower() in [d.lower() for d in date_candidates]:
                date_col = c
            if value_col is None and c.strip().lower() in [v.lower() for v in value_candidates]:
                value_col = c

        if date_col is None or value_col is None:
            raise KeyError(f"Colunas necessárias não encontradas. Encontradas: {cols}")

        # Parse de data e valores
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df[value_col] = pd.to_numeric(df[value_col].astype(str).str.replace(',',''), errors='coerce')

        mask = (df[date_col].dt.year == 2026) & (df[date_col].dt.month == 1)
        total = df.loc[mask, value_col].sum(skipna=True)

        print(f"Arquivo usado: {META_CSV}")
        print(f"Coluna de data: {date_col}")
        print(f"Coluna de valor: {value_col}")
        print(f"O investimento de janeiro de 2026 em Meta foi de R$ {total:,.2f}")
    except Exception as e:
        print(f"Erro ao calcular o investimento: {e}")


if __name__ == "__main__":
    investimento_janeiro_2026_meta()
