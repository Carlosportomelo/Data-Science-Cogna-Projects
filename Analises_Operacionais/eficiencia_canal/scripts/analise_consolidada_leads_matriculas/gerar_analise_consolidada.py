import pandas as pd
from pathlib import Path
from datetime import datetime

def find_hubspot_file():
    repo_root = Path(__file__).resolve().parents[3]
    candidates = [
        Path(r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_negocios_perdidos_ATUAL.csv"),
        Path(r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_ATUAL.csv"),
        repo_root / 'data' / 'hubspot_leads.csv',
        Path.cwd() / 'data' / 'hubspot_leads.csv',
    ]
    for p in candidates:
        if p and p.exists():
            return p
    raise FileNotFoundError(f"Nenhum arquivo hubspot encontrado. Procurado: {candidates}")

def load_hubspot(p):
    encs = ['utf-8', 'latin-1', 'cp1252']
    for e in encs:
        try:
            return pd.read_csv(p, encoding=e, low_memory=False)
        except Exception:
            continue
    raise ValueError('Não foi possível ler o CSV do HubSpot com encodings testados')

def detect_cols(df):
    cols = {c.lower(): c for c in df.columns}
    col_create = next((v for k,v in cols.items() if any(x in k for x in ['data de criação','data de criacao','create date','created'])), None)
    col_close = next((v for k,v in cols.items() if any(x in k for x in ['data de fechamento','close date','closedate','close_date'])), None)
    col_unit = next((v for k,v in cols.items() if any(x in k for x in ['unidade desejada','unidade','campus','business unit'])), None)
    col_stage = next((v for k,v in cols.items() if 'etapa' in k or 'stage' in k or 'status' in k), None)
    col_source = next((v for k,v in cols.items() if 'fonte' in k or 'origem' in k or 'source' in k), None)
    return col_create, col_close, col_unit, col_stage, col_source

def build_cycles():
    # Ciclos: formato 'YY.1' -> Início: 01/out/(YY-1)  - Fim: 20/02/(YY)
    cycles = {}
    for yy in ['23', '24', '25', '26']:
        label = f"{yy}.1"
        start_year = 2000 + int(yy) - 1
        start = pd.Timestamp(f"{start_year}-10-01")
        end = pd.Timestamp(f"{start_year + 1}-02-20") + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        cycles[label] = (start, end)
    return cycles

def generate_reports(df, col_create, col_close, col_unit, col_stage, col_source):
    cycles = build_cycles()

    # parse dates
    if col_create:
        df['_dt_create'] = pd.to_datetime(df[col_create], errors='coerce')
    else:
        df['_dt_create'] = pd.NaT
    if col_close:
        df['_dt_close'] = pd.to_datetime(df[col_close], errors='coerce')
    else:
        df['_dt_close'] = pd.NaT

    # Resumo por ciclo
    rows = []
    for ciclo, (start, end) in cycles.items():
        leads_mask = df['_dt_create'].between(start, end)
        leads_count = int(leads_mask.sum())

        matr_mask = df[col_stage].astype(str).str.contains('MATRÍCULA CONCLUÍDA', case=False, na=False) if col_stage else pd.Series([False]*len(df))
        matr_mask = matr_mask & df['_dt_close'].between(start, end)
        matr_count = int(matr_mask.sum())

        rows.append({'Ciclo': ciclo, 'Leads': leads_count, 'Matriculas': matr_count})

    df_summary = pd.DataFrame(rows)

    # Leads (Unidades)
    unit_col = col_unit if col_unit else None
    if unit_col:
        leads_units = {}
        matr_units = {}
        for ciclo, (start, end) in cycles.items():
            leads_units[ciclo] = df[df['_dt_create'].between(start, end)].groupby(unit_col).size()
            matr_units[ciclo] = df[df['_dt_close'].between(start, end) & df[col_stage].astype(str).str.contains('MATRÍCULA CONCLUÍDA', case=False, na=False)].groupby(unit_col).size()

        df_leads_units = pd.DataFrame(leads_units).fillna(0).astype(int)
        df_leads_units.index.name = 'Unidade'

        df_matr_units = pd.DataFrame(matr_units).fillna(0).astype(int)
        df_matr_units.index.name = 'Unidade'
    else:
        df_leads_units = pd.DataFrame()
        df_matr_units = pd.DataFrame()

    # Leads / Matrículas por Canal (Fonte)
    if col_source:
        leads_chan = {}
        matr_chan = {}
        for ciclo, (start, end) in cycles.items():
            leads_chan[ciclo] = df[df['_dt_create'].between(start, end)].groupby(col_source).size()
            matr_chan[ciclo] = df[df['_dt_close'].between(start, end) & df[col_stage].astype(str).str.contains('MATRÍCULA CONCLUÍDA', case=False, na=False)].groupby(col_source).size()
        df_leads_chan = pd.DataFrame(leads_chan).fillna(0).astype(int)
        df_matr_chan = pd.DataFrame(matr_chan).fillna(0).astype(int)
    else:
        df_leads_chan = pd.DataFrame()
        df_matr_chan = pd.DataFrame()

    return df_summary, df_leads_units, df_matr_units, df_leads_chan, df_matr_chan

def main():
    print('Gerando análise consolidada...')
    p = find_hubspot_file()
    print('Lendo:', p)
    df = load_hubspot(p)

    col_create, col_close, col_unit, col_stage, col_source = detect_cols(df)
    if col_stage is None:
        # fallback common name
        col_stage = next((c for c in df.columns if 'etapa' in c.lower()), None)

    summary, leads_units, matr_units, leads_chan, matr_chan = generate_reports(df, col_create, col_close, col_unit, col_stage, col_source)

    # Escreve Excel
    out_dir = Path(__file__).resolve().parents[3] / 'outputs' / 'analise_consolidada_leads_matriculas'
    out_dir.mkdir(parents=True, exist_ok=True)
    date_str = datetime.now().strftime('%Y-%m-%d')
    out_file = out_dir / f'analise_consolidada_leads_matriculas_{date_str}.xlsx'

    try:
        with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
            summary.to_excel(writer, sheet_name='Resumo e Insights', index=False)
            if not leads_units.empty:
                leads_units.to_excel(writer, sheet_name='Leads (Unidades)')
            if not matr_units.empty:
                matr_units.to_excel(writer, sheet_name='Matrículas (Unidades)')
            if not leads_chan.empty:
                leads_chan.to_excel(writer, sheet_name='Leads (Canais)')
            if not matr_chan.empty:
                matr_chan.to_excel(writer, sheet_name='Matrículas (Canais)')
    except PermissionError:
        # Arquivo pode estar aberto no Excel — tenta nome alternativo com timestamp
        alt_file = out_dir / f'analise_consolidada_leads_matriculas_{date_str}_{datetime.now().strftime("%H%M%S")}.xlsx'
        print(f'Permissão negada ao escrever {out_file}. Tentando gravar em {alt_file} ...')
        try:
            with pd.ExcelWriter(alt_file, engine='openpyxl') as writer:
                summary.to_excel(writer, sheet_name='Resumo e Insights', index=False)
                if not leads_units.empty:
                    leads_units.to_excel(writer, sheet_name='Leads (Unidades)')
                if not matr_units.empty:
                    matr_units.to_excel(writer, sheet_name='Matrículas (Unidades)')
                if not leads_chan.empty:
                    leads_chan.to_excel(writer, sheet_name='Leads (Canais)')
                if not matr_chan.empty:
                    matr_chan.to_excel(writer, sheet_name='Matrículas (Canais)')
            out_file = alt_file
        except Exception as e:
            print('Falha ao gravar arquivo alternativo:', e)
            raise

    print('Relatório gerado em:', out_file)

if __name__ == '__main__':
    main()
