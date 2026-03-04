#!/usr/bin/env python3
"""
Codifica motivos em números (mascara) e grava glossário.

Entrada preferencial: Outputs/hubspot_negocios_perdidos_id_motivo_sinonimos_*.csv
Saída:
 - Outputs/hubspot_negocios_perdidos_id_motivo_codificado_<date>.csv (Record ID + Motivo_Code)
 - Outputs/glossario_motivos_<date>.csv (Motivo_Code -> Motivo_Normalizado)

Imprime o glossário no terminal.
"""
import os
import glob
import pandas as pd
from datetime import datetime


def find_latest(path_pattern_list):
    for p in path_pattern_list:
        candidates = glob.glob(p)
        if candidates:
            candidates.sort(key=os.path.getmtime, reverse=True)
            return candidates[0]
    return None


def main():
    candidates = [
        os.path.join('Outputs', 'hubspot_negocios_perdidos_id_motivo_sinonimos_*.csv'),
        os.path.join('Outputs', 'hubspot_negocios_perdidos_id_motivo_*.csv'),
        os.path.join('Outputs', 'sanitized_preview_sample.csv'),
        os.path.join('Data', 'hubspot_negocios_perdidos.csv')
    ]
    inp = find_latest(candidates)
    if inp is None:
        print('Nenhum arquivo fonte encontrado para codificação.')
        return

    print('Lendo arquivo fonte:', inp)
    df = pd.read_csv(inp, dtype=str, low_memory=False)

    # Detect Record ID column
    id_col = None
    for c in df.columns:
        if c.lower().strip() in ('record id', 'recordid', 'record_id', 'id'):
            id_col = c
            break
    if id_col is None:
        for c in df.columns:
            if 'id' in c.lower():
                id_col = c
                break
    if id_col is None:
        print('Coluna ID não encontrada.')
        return

    # Detect reason column (normalized or original)
    reason_col = None
    for name in ['Motivo_Normalizado', 'Motivo do negócio perdido', 'Motivo do negocio perdido', 'Motivo']:
        for c in df.columns:
            if c == name:
                reason_col = c
                break
        if reason_col:
            break
    if reason_col is None:
        for c in df.columns:
            if 'motivo' in c.lower():
                reason_col = c
                break
    if reason_col is None:
        print('Coluna de motivo não encontrada.')
        return

    # build mapping
    uniq = pd.Series(df[reason_col].fillna('Outro motivo').astype(str).unique())
    uniq = uniq.sort_values().reset_index(drop=True)
    mapping = {v: i+1 for i, v in enumerate(uniq)}

    # create encoded df with masked column
    out_df = pd.DataFrame()
    out_df['Record ID'] = df[id_col].astype(str)
    out_df['Motivo_Code'] = df[reason_col].fillna('Outro motivo').astype(str).map(mapping)

    date = datetime.now().strftime('%Y-%m-%d')
    os.makedirs('Outputs', exist_ok=True)
    out_path = os.path.join('Outputs', f'hubspot_negocios_perdidos_id_motivo_codificado_{date}.csv')
    glossary_path = os.path.join('Outputs', f'glossario_motivos_{date}.csv')

    out_df.to_csv(out_path, index=False)
    pd.DataFrame([{'Motivo_Code': mapping[k], 'Motivo_Normalizado': k} for k in mapping]).sort_values('Motivo_Code').to_csv(glossary_path, index=False)

    # print glossary to terminal
    print('\nGlossário (Motivo_Code -> Motivo_Normalizado):')
    for k, v in sorted(mapping.items(), key=lambda x: x[1]):
        print(f"{v}: {k}")

    print('\nArquivo codificado salvo em:', out_path)
    print('Glossário salvo em:', glossary_path)
    print('\nAmostra do arquivo codificado (10 linhas):')
    print(out_df.head(10).to_string(index=False))


if __name__ == '__main__':
    main()
