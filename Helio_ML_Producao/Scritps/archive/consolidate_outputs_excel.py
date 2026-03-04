#!/usr/bin/env python3
"""
Consolidate multiple Outputs CSV/XLSX into a single Excel workbook with separate sheets.

Sheets included (if files found):
- Top50, Top250, Top1000, Top10_per_unit, Qualification_All
- Performance plan: Daily_Plan, Weekly_Plan
- Looker dataset (one sheet)
- Operational pack (if exists, its sheets will be copied)

Output:
  Outputs/Executive_Package_AllSheets_<date>.xlsx

Usage:
  python Scritps/consolidate_outputs_excel.py
"""
import os
import glob
from datetime import datetime
import pandas as pd


def latest(path_pattern):
    files = glob.glob(path_pattern, recursive=True)
    if not files:
        return None
    files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return files[0]


def try_read_csv(path):
    try:
        return pd.read_csv(path, low_memory=False)
    except Exception:
        try:
            return pd.read_excel(path)
        except Exception:
            return None


def main():
    os.makedirs('Outputs', exist_ok=True)
    stamp = datetime.now().strftime('%Y-%m-%d')
    out_xlsx = f'Outputs/Pacote_Executivo_TodasAbas_{stamp}.xlsx'

    # Find top lists
    top50 = latest('Outputs/**/top50_*.csv')
    top250 = latest('Outputs/**/top250_*.csv')
    top1000 = latest('Outputs/**/top1000_*.csv')
    top10_unit = latest('Outputs/**/top10_per_unit_*.csv')
    qual_all = (latest('Outputs/**/qualification_all_scored_*.csv') or latest('Outputs/**/qualification_all_scored_*.xlsx') 
                or latest('Outputs/**/qualificacao_todos_pontuados_*.csv') or latest('Outputs/**/qualificacao_todos_pontuados_*.xlsx')
                or latest('Outputs/**/performance_plan_*_qualification_scored.csv'))

    perf_daily = latest('Outputs/**/performance_plan_*_daily.csv')
    perf_weekly = latest('Outputs/**/performance_plan_*_weekly.csv')

    looker = latest('Outputs/**/looker_dataset*.csv') or latest('Outputs/**/looker_dataset*.xlsx')
    operational_pack = latest('Outputs/**/Operational_Pack_Executive_*.xlsx') or latest('Outputs/**/Operational_Pack_Executive_*.xls')

    sheets = []
    # read and collect
    if top50:
        df = try_read_csv(top50)
        if df is not None:
            sheets.append(('Top50', df))
    if top250:
        df = try_read_csv(top250)
        if df is not None:
            sheets.append(('Top250', df))
    if top1000:
        df = try_read_csv(top1000)
        if df is not None:
            sheets.append(('Top1000', df))
    if top10_unit:
        df = try_read_csv(top10_unit)
        if df is not None:
            sheets.append(('Top10_por_unidade', df))
    if qual_all:
        df = try_read_csv(qual_all)
        if df is not None:
            sheets.append(('Qualificacao_Todos', df))

    if perf_daily:
        df = try_read_csv(perf_daily)
        if df is not None:
            sheets.append(('Plano_Diario', df))
    if perf_weekly:
        df = try_read_csv(perf_weekly)
        if df is not None:
            sheets.append(('Plano_Semanal', df))

    if looker:
        df = try_read_csv(looker)
        if df is not None:
            sheets.append(('Dataset_Looker', df))

    # If there is an operational pack xlsx, try to copy its sheets
    if operational_pack:
        try:
            op = pd.read_excel(operational_pack, sheet_name=None)
            for name, df in op.items():
                # avoid duplicate sheet names
                sheet_name = f'Op_{name}' if name in [s[0] for s in sheets] else name
                sheets.append((sheet_name, df))
        except Exception:
            pass

    if not sheets:
        print('No known outputs found to consolidate. Check Outputs/ folder.')
        return

    # write consolidated workbook
    with pd.ExcelWriter(out_xlsx, engine='openpyxl') as w:
        for name, df in sheets:
            # sanitize sheet name
            sheet = name[:31]
            try:
                df.to_excel(w, sheet_name=sheet, index=False)
            except Exception:
                # if write fails, try to coerce all columns to strings
                df2 = df.astype(str)
                df2.to_excel(w, sheet_name=sheet, index=False)

    print('Consolidated workbook saved:', out_xlsx)


if __name__ == '__main__':
    main()
