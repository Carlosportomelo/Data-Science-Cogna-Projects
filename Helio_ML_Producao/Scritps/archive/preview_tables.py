#!/usr/bin/env python3
"""
Exibe amostras tabulares das bases bruta e sanitizada de negócios perdidos.

Saída:
- Imprime no terminal uma tabela limpa (10 linhas) com colunas chave.
- Salva previews em `Outputs/raw_preview_sample.csv` e `Outputs/sanitized_preview_sample.csv`.

Uso:
  python Scritps/preview_tables.py
"""
import os
import glob
import pandas as pd


def truncate_series(s, width=50):
    if pd.isna(s):
        return ''
    s = str(s)
    return s if len(s) <= width else s[:width-30] + ' ... ' + s[-30:]


def prepare(df, cols, width=50):
    present = [c for c in cols if c in df.columns]
    out = df[present].copy()
    for c in out.columns:
        if out[c].dtype == object:
            out[c] = out[c].apply(lambda v: truncate_series(v, width))
    return out


def main():
    os.makedirs('Outputs', exist_ok=True)

    raw_path = os.path.join('Data', 'hubspot_negocios_perdidos.csv')
    san_files = glob.glob(os.path.join('Outputs', 'archive', 'hubspot_negocios_perdidos_sanitizado_*.csv'))

    if not os.path.exists(raw_path):
        print('Arquivo bruto não encontrado:', raw_path)
        return

    if not san_files:
        print('Arquivo sanitizado não encontrado em Outputs/archive')
        return

    san_files.sort(key=os.path.getmtime, reverse=True)
    san_path = san_files[0]

    df_raw = pd.read_csv(raw_path, dtype=str, low_memory=False)
    df_san = pd.read_csv(san_path, dtype=str, low_memory=False)

    cols = ['Record ID', 'record_hash', 'Motivo do negócio perdido', 'Etapa do negócio', 'Unidade Desejada', 'Número de atividades de vendas', 'Data de criação', 'Valor na moeda da empresa']

    sample_raw = prepare(df_raw.head(10), cols, width=60)
    sample_san = prepare(df_san.head(10), cols, width=60)

    # print nice tables
    pd.set_option('display.width', 200)
    pd.set_option('display.max_colwidth', 60)
    print('\n=== AMOSTRA - BASE BRUTA (10 linhas) ===\n')
    print(sample_raw.to_string(index=False))
    print('\n=== AMOSTRA - BASE SANITIZADA (10 linhas) ===\n')
    print(sample_san.to_string(index=False))

    # save previews
    raw_preview = os.path.join('Outputs', 'raw_preview_sample.csv')
    san_preview = os.path.join('Outputs', 'sanitized_preview_sample.csv')
    sample_raw.to_csv(raw_preview, index=False)
    sample_san.to_csv(san_preview, index=False)
    print(f'\nPreviews salvos: {raw_preview}, {san_preview}')


if __name__ == '__main__':
    main()
