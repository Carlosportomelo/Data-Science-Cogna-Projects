import sys
import pandas as pd
from pathlib import Path

def compare_sheets(file_a, file_b):
    xa = pd.ExcelFile(file_a)
    xb = pd.ExcelFile(file_b)
    sheets_a = xa.sheet_names
    sheets_b = xb.sheet_names
    print(f"Comparing:\n A: {file_a}\n B: {file_b}\n")
    print("Sheets A:", sheets_a)
    print("Sheets B:", sheets_b)
    common = [s for s in sheets_a if s in sheets_b]
    only_a = [s for s in sheets_a if s not in sheets_b]
    only_b = [s for s in sheets_b if s not in sheets_a]
    if only_a:
        print("Only in A:", only_a)
    if only_b:
        print("Only in B:", only_b)
    print('\n--- Comparing common sheets ---')
    for s in common:
        da = xa.parse(s)
        db = xb.parse(s)
        print(f"\nSheet: {s}")
        print(f" A shape: {da.shape} | B shape: {db.shape}")
        cols_a = list(da.columns)
        cols_b = list(db.columns)
        if cols_a != cols_b:
            print("  Columns differ")
            print("   A columns:", cols_a)
            print("   B columns:", cols_b)
        # Compare head rows
        n = min(5, max(len(da), len(db)))
        ha = da.head(n).fillna('<<NA>>').astype(str)
        hb = db.head(n).fillna('<<NA>>').astype(str)
        if ha.equals(hb):
            print("  First rows equal")
        else:
            print("  First rows differ. Showing diff by cell:")
            # align columns
            allcols = sorted(set(ha.columns).union(hb.columns))
            ha2 = ha.reindex(columns=allcols, fill_value='<<MISSING>>')
            hb2 = hb.reindex(columns=allcols, fill_value='<<MISSING>>')
            for i in range(n):
                row_a = ha2.iloc[i]
                row_b = hb2.iloc[i]
                diffs = [c for c in allcols if str(row_a.get(c,'')).strip() != str(row_b.get(c,'')).strip()]
                if diffs:
                    print(f"   Row {i+1} diffs columns: {diffs}")
                    for c in diffs:
                        print(f"    {c}: A='{row_a.get(c)}'  B='{row_b.get(c)}'")

if __name__ == '__main__':
    PATH = Path(__file__).resolve().parents[1] / 'outputs'
    # originals (no timestamp) vs newest timestamped (find newest by modified time)
    files = list(PATH.glob('Analise_Ciclo_26_1*.xlsx'))
    # Separate originals (exact names) and timestamped
    orig_main = PATH / 'Analise_Ciclo_26_1.xlsx'
    orig_un = PATH / 'Analise_Ciclo_26_1_POR_UNIDADE.xlsx'
    orig_err = PATH / 'Analise_Ciclo_26_1_ERROS_CRM.xlsx'

    def newest_matching(prefix):
        candidates = sorted(PATH.glob(prefix + '_*.xlsx'), key=lambda p: p.stat().st_mtime, reverse=True)
        return candidates[0] if candidates else None

    new_main = newest_matching('Analise_Ciclo_26_1')
    new_un = newest_matching('Analise_Ciclo_26_1_POR_UNIDADE')
    new_err = newest_matching('Analise_Ciclo_26_1_ERROS_CRM')

    pairs = [
        (orig_main, new_main),
        (orig_un, new_un),
        (orig_err, new_err)
    ]
    for a,b in pairs:
        if not a.exists():
            print(f"Original missing: {a}")
            continue
        if b is None or not b.exists():
            print(f"New file missing for prefix of {a.name}")
            continue
        compare_sheets(str(a), str(b))
