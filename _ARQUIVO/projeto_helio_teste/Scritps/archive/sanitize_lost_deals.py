#!/usr/bin/env python3
"""
Sanitiza a base de negócios perdidos removendo PII e preservando colunas úteis para ML.

Comportamento:
- Lê `Data/hubspot_negocios_perdidos.csv` por padrão (ou caminho especificado).
- Remove/exclui colunas sensíveis (nome, email, telefone, cpf, endereço, proprietário etc.).
- Gera hash consistente (SHA256 + salt) para IDs disponíveis em coluna que contenha 'id'.
- Mantém datas, valores, contadores e colunas categóricas úteis.
- Grava resultado em `Outputs/archive/hubspot_negocios_perdidos_sanitizado_<date>.csv`.

Uso:
  python Scritps/sanitize_lost_deals.py
  python Scritps/sanitize_lost_deals.py --input Data/hubspot_negocios_perdidos.csv --output Outputs/archive/sanitized.csv

Observação:
Para produção, defina a variável de ambiente `SANITIZE_SALT` para alterar o salt do hash.
"""
import os
import argparse
from datetime import datetime
import hashlib
import pandas as pd
import numpy as np


def hash_value(v, salt):
    if pd.isna(v):
        return ''
    s = f"{salt}|{v}"
    return hashlib.sha256(s.encode('utf-8')).hexdigest()[:16]


def detect_id_column(df):
    for c in df.columns:
        low = c.lower()
        if low in ('record id', 'record_id', 'id', 'lead id', 'lead_id') or low.endswith('id') or 'hs_object' in low:
            return c
    return None


def sanitize(df, salt):
    # columns to drop by pattern (include common Portuguese variants)
    drop_patterns = ['nome', 'name', 'email', 'phone', 'telefone', 'cel', 'cpf', 'rg', 'document', 'endereço', 'endereco', 'cep', 'bairro', 'rua', 'address', 'owner', 'propriet', 'proprietário', 'proprietario']
    cols_to_drop = []
    for c in df.columns:
        cl = c.lower()
        for p in drop_patterns:
            if p in cl:
                cols_to_drop.append(c)
                break

    # drop obvious PII
    cols_to_drop = list(dict.fromkeys(cols_to_drop))

    # keep a copy of dropped column names for log
    dropped = [c for c in cols_to_drop if c in df.columns]

    df2 = df.copy()
    df2.drop(columns=dropped, inplace=True, errors='ignore')

    # find id column and create hashed id
    id_col = detect_id_column(df)
    if id_col is None:
        # try common lead id
        for c in df.columns:
            if 'id' in c.lower():
                id_col = c
                break

    if id_col is not None:
        df2['record_hash'] = df[id_col].astype(str).apply(lambda v: hash_value(v, salt))
        # remove the original id column to avoid leaking identifiers
        if id_col in df2.columns:
            df2.drop(columns=[id_col], inplace=True, errors='ignore')
    else:
        # create synthetic hashed id from row index
        df2['record_hash'] = [hash_value(i, salt) for i in range(len(df2))]

    # Ensure date columns are parsed and keep Year/Month
    for c in df2.columns:
        if 'date' in c.lower() or 'data' in c.lower():
            try:
                df2[c] = pd.to_datetime(df2[c], errors='coerce')
                # add year/month for ML
                df2[f'{c}_year'] = df2[c].dt.year
                df2[f'{c}_month'] = df2[c].dt.month
            except Exception:
                pass

    # Fill NA for numeric columns with sentinel (-1) and for categorical with 'missing'
    for c in df2.columns:
        if pd.api.types.is_numeric_dtype(df2[c]):
            df2[c] = df2[c].fillna(-1)
        else:
            df2[c] = df2[c].fillna('missing')

    # reorder: put record_hash first
    cols = ['record_hash'] + [c for c in df2.columns if c != 'record_hash']
    df2 = df2[cols]

    meta = {
        'dropped_columns': dropped,
        'id_column_used': id_col
    }

    return df2, meta


def main():
    parser = argparse.ArgumentParser(description='Sanitize lost deals dataset (remove PII)')
    parser.add_argument('--input', '-i', default=os.path.join('Data', 'hubspot_negocios_perdidos.csv'))
    parser.add_argument('--output', '-o', default=None)
    parser.add_argument('--salt', '-s', default=os.environ.get('SANITIZE_SALT', 'change_this_salt'))
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print('Arquivo de input não encontrado:', args.input)
        return

    df = pd.read_csv(args.input, dtype=str, low_memory=False)
    print(f'Linhas lidas: {len(df):,}')

    sanitized, meta = sanitize(df, args.salt)

    date = datetime.now().strftime('%Y-%m-%d')
    os.makedirs(os.path.join('Outputs', 'archive'), exist_ok=True)
    output = args.output or os.path.join('Outputs', 'archive', f'hubspot_negocios_perdidos_sanitizado_{date}.csv')

    sanitized.to_csv(output, index=False)
    print('Arquivo sanitizado gravado em:', output)
    print('Colunas removidas:', meta['dropped_columns'])
    print('ID usado para hash:', meta['id_column_used'])


if __name__ == '__main__':
    main()
