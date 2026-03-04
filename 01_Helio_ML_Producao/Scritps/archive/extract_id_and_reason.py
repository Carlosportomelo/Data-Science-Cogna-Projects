#!/usr/bin/env python3
"""
Extrai apenas as colunas `Record ID` e `Motivo do negócio perdido` da base de negócios perdidos.

Gera: Outputs/hubspot_negocios_perdidos_id_motivo_<date>.csv
"""
import os
import pandas as pd
from datetime import datetime


def find_column(df, candidates):
    for c in df.columns:
        if c in candidates:
            return c
        lc = c.lower().strip()
        for cand in candidates:
            if lc == cand.lower():
                return c
    return None


def main():
    inp = os.path.join('Data', 'hubspot_negocios_perdidos.csv')
    if not os.path.exists(inp):
        print('Arquivo de entrada não encontrado:', inp)
        return

    df = pd.read_csv(inp, dtype=str, low_memory=False)

    id_col = find_column(df, ['Record ID', 'RecordID', 'record id', 'id'])
    reason_col = find_column(df, ['Motivo do negócio perdido', 'Motivo do negocio perdido', 'motivo do negócio perdido', 'motivo do negocio perdido', 'Motivo'])

    if id_col is None:
        print('Coluna de ID não encontrada.')
        return
    if reason_col is None:
        print('Coluna de motivo não encontrada.')
        return

    out = df[[id_col, reason_col]].copy()
    out.columns = ['Record ID', 'Motivo do negócio perdido']

    date = datetime.now().strftime('%Y-%m-%d')
    os.makedirs('Outputs', exist_ok=True)
    out_path = os.path.join('Outputs', f'hubspot_negocios_perdidos_id_motivo_{date}.csv')
    out.to_csv(out_path, index=False)
    print('Arquivo gerado:', out_path)
    print('\nAmostra (10 linhas):')
    print(out.head(10).to_string(index=False))


if __name__ == '__main__':
    main()
