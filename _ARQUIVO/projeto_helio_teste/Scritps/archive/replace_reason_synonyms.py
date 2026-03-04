#!/usr/bin/env python3
"""
Substitui textos da coluna `Motivo do negócio perdido` por sinônimos/normalizações.

Regra: aplicação de mapeamento por substring para reduzir variação textual e sensibilidade.
Leitura: prioriza `Outputs/hubspot_negocios_perdidos_id_motivo_*.csv` mais recente.
Saída: `Outputs/hubspot_negocios_perdidos_id_motivo_sinonimos_<date>.csv`.
"""
import os
import glob
import pandas as pd
from datetime import datetime


def normalize_reason(text):
    if pd.isna(text):
        return 'Outro motivo'
    s = str(text).strip()
    sl = s.lower()
    if 'idade' in sl:
        return 'Incompatibilidade de idade'
    if 'logíst' in sl or 'transp' in sl or 'transporte' in sl:
        return 'Problema logístico'
    if 'fora do perfil' in sl or 'fora do perfil' in sl or 'fora do perfil' in sl:
        return 'Fora do perfil'
    if 'finance' in sl or 'pagam' in sl or 'dinheiro' in sl:
        return 'Motivo financeiro'
    if 'teste' in sl or 'teste marketing' in sl:
        return 'Fora do perfil'
    if 'desist' in sl or 'cancel' in sl:
        return 'Desistência'
    if 'sem interesse' in sl or 'nao interessado' in sl or 'não interessado' in sl:
        return 'Sem interesse'
    if 'nao respondeu' in sl or 'não respondeu' in sl or 'sem resposta' in sl:
        return 'Sem resposta'
    if 'document' in sl or 'documenta' in sl:
        return 'Problema documental'
    if 'logistica' in sl:
        return 'Problema logístico'
    # common Portuguese variants
    if 'finan' in sl:
        return 'Motivo financeiro'
    # fallback: keep short normalized version (first 60 chars) or generic
    return 'Outro motivo'


def find_latest_inputs():
    candidates = glob.glob(os.path.join('Outputs', 'hubspot_negocios_perdidos_id_motivo_*.csv'))
    if candidates:
        candidates.sort(key=os.path.getmtime, reverse=True)
        return candidates[0]
    # fallback to sanitized archive
    candidates = glob.glob(os.path.join('Outputs', 'archive', 'hubspot_negocios_perdidos_sanitizado_*.csv'))
    if candidates:
        candidates.sort(key=os.path.getmtime, reverse=True)
        return candidates[0]
    # fallback to raw data
    raw = os.path.join('Data', 'hubspot_negocios_perdidos.csv')
    if os.path.exists(raw):
        return raw
    return None


def main():
    inp = find_latest_inputs()
    if inp is None:
        print('Nenhum arquivo de entrada encontrado (procure em Outputs/ ou Data/)')
        return

    print('Lendo:', inp)
    df = pd.read_csv(inp, dtype=str, low_memory=False)

    # try to detect reason column
    reason_col = None
    for c in df.columns:
        if c.lower().strip() in ('motivo do negócio perdido', 'motivo do negocio perdido', 'motivo'):
            reason_col = c
            break
    if reason_col is None:
        # try fuzzy
        for c in df.columns:
            if 'motivo' in c.lower():
                reason_col = c
                break

    if reason_col is None:
        print('Coluna de motivo não encontrada no arquivo:', inp)
        return

    df_out = df.copy()
    df_out['Motivo_Normalizado'] = df_out[reason_col].apply(normalize_reason)

    date = datetime.now().strftime('%Y-%m-%d')
    out_dir = 'Outputs'
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f'hubspot_negocios_perdidos_id_motivo_sinonimos_{date}.csv')
    df_out.to_csv(out_path, index=False)

    print('Arquivo com motivos normalizados gravado em:', out_path)
    print('\nAmostra (10 linhas):')
    print(df_out[[c for c in df_out.columns if c.lower().strip() in ("record id","recordid","record_id")]+[reason_col,'Motivo_Normalizado']].head(10).to_string(index=False))


if __name__ == '__main__':
    main()
