from pathlib import Path
import pandas as pd
import unicodedata
import numpy as np
from difflib import get_close_matches


def normalize(s: str) -> str:
    if pd.isna(s):
        return ""
    s = str(s)
    s = s.strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = " ".join(s.split())
    return s


EXCEL_PATH = Path(r"C:\Users\a483650\Projetos\03_Analises_Operacionais\eficiencia_canal\outputs\analise_consolidada_leads_matriculas.xlsx")
CSV_PATH = Path(r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_atual.csv")

# Units to check (user-provided spellings) - normalized when matching
UNITS = ["allphaville", "aldeota", "agua fira"]

# Optional manual aliases (common variants) to help matching
ALIASES = {
    "allphaville": ["alphaville", "allphaville", "alphaville -"],
    "aldeota": ["aldeota"],
    "agua fira": ["agua fria", "agua fira", "água fria", "agua-fria"],
}

# Ciclo 26.1 date range (used to filter CSV by creation date)
START_26 = pd.to_datetime("2025-10-01")
END_26 = pd.to_datetime("2026-02-20")

ETAPA_MATRICULA = "MATRÍCULA CONCLUÍDA"


def read_csv_attempt(path: Path) -> pd.DataFrame:
    for sep in [",", ";"]:
        try:
            df = pd.read_csv(path, sep=sep, low_memory=False)
            if len(df.columns) > 1:
                return df
        except Exception:
            continue
    # last resort
    return pd.read_csv(path, sep=",", low_memory=False)


def find_column(df: pd.DataFrame, candidates):
    cols = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    # try approximate match
    for k, v in cols.items():
        for cand in candidates:
            if cand.lower() in k:
                return v
    return None


def csv_counts(df: pd.DataFrame, units: list) -> dict:
    # Normalize column names and find relevant columns
    cols = {c.lower().strip(): c for c in df.columns}
    # possible column names
    unidade_col = find_column(df, ["Unidade", "Unidade Desejada", "unidade desejada", "campus", "unidade"])
    etapa_col = find_column(df, ["Etapa do negócio", "Stage", "Stage name", "etapa do negocio"])
    data_col = find_column(df, ["Data de criação", "Create Date", "Created Date", "Create date"])

    if unidade_col:
        df[unidade_col] = df[unidade_col].astype(str).apply(normalize)

    if etapa_col:
        df["is_matricula"] = df[etapa_col].astype(str).str.contains(ETAPA_MATRICULA, case=False, na=False)
    else:
        df["is_matricula"] = False

    if data_col:
        df[data_col] = pd.to_datetime(df[data_col], errors="coerce")
        mask_26 = (df[data_col] >= START_26) & (df[data_col] <= END_26)
    else:
        mask_26 = pd.Series([False] * len(df))

    results = {}
    for u in units:
        nu = normalize(u)
        if unidade_col:
            sub_all = df[df[unidade_col].str.contains(nu, na=False)]
            # se não encontrou, tente aliases e correspondência aproximada
            if sub_all.empty:
                # aliases
                al = ALIASES.get(u, [])
                for a in al:
                    sub_all = df[df[unidade_col].str.contains(normalize(a), na=False)]
                    if not sub_all.empty:
                        print(f"CSV: Unidade '{u}' mapeada via alias '{a}'")
                        break
            if sub_all.empty:
                # tentativa por correspondência aproximada
                unique_vals = df[unidade_col].dropna().astype(str).apply(normalize).unique().tolist()
                matches = get_close_matches(nu, unique_vals, n=3, cutoff=0.8)
                if matches:
                    match = matches[0]
                    sub_all = df[df[unidade_col].astype(str).apply(lambda x: normalize(x) == match)]
                    if not sub_all.empty:
                        print(f"CSV: Unidade '{u}' mapeada via fuzzy match '{match}'")
        else:
            sub_all = df[df.apply(lambda r: nu in " ".join(map(str, r.values)).lower(), axis=1)]

        leads_total = len(sub_all)
        mat_total = int(sub_all["is_matricula"].sum())

        sub_26 = sub_all[mask_26.loc[sub_all.index]] if data_col else pd.DataFrame(columns=df.columns)
        leads_26 = len(sub_26)
        mat_26 = int(sub_26["is_matricula"].sum()) if not sub_26.empty else 0

        results[u] = {
            "leads_total_csv": leads_total,
            "mat_total_csv": mat_total,
            "leads_26_csv": leads_26,
            "mat_26_csv": mat_26,
        }

    return results


def excel_extract_units(path: Path, units: list) -> dict:
    res = {u: {"leads_excel": None, "mat_excel": None, "sheet_leads": None, "sheet_mat": None} for u in units}
    try:
        xls = pd.read_excel(path, sheet_name=None)
    except Exception as e:
        print(f"Erro ao abrir Excel: {e}")
        return res

    # heuristics: find sheets that contain a 'Unidade' column and numeric columns
    for name, df in xls.items():
        cols = [c for c in df.columns]
        if any("unidade" in str(c).lower() for c in cols):
            # try to find a numeric column for current cycle or TOTAL GERAL
            candidate_cols = [c for c in cols if isinstance(c, str) and ("total" in c.lower() or "26.1" in c or "leads" in c.lower() or "matric" in c.lower())]
            if not candidate_cols:
                # try any numeric column
                numeric = [c for c in cols if pd.api.types.is_numeric_dtype(df[c])]
                candidate_cols = numeric

            if not candidate_cols:
                continue

            unidade_col = [c for c in cols if "unidade" in str(c).lower()][0]
            # normalize unidade values in this sheet
            df["__unidade_norm"] = df[unidade_col].astype(str).apply(normalize)

            for u in units:
                nu = normalize(u)
                row = df[df["__unidade_norm"].str.contains(nu, na=False)]
                if row.empty:
                    # aliases
                    al = ALIASES.get(u, [])
                    for a in al:
                        row = df[df["__unidade_norm"].str.contains(normalize(a), na=False)]
                        if not row.empty:
                            print(f"Excel: Unidade '{u}' mapeada via alias '{a}' na sheet '{name}'")
                            break
                if row.empty:
                    # fuzzy
                    unique_vals = df["__unidade_norm"].dropna().unique().tolist()
                    matches = get_close_matches(nu, unique_vals, n=3, cutoff=0.8)
                    if matches:
                        mv = matches[0]
                        row = df[df["__unidade_norm"] == mv]
                        if not row.empty:
                            print(f"Excel: Unidade '{u}' mapeada via fuzzy match '{mv}' na sheet '{name}'")
                if row.empty:
                    continue
                # prefer TOTAL GERAL or 26.1 if present
                chosen = None
                for cand in [c for c in candidate_cols if isinstance(c, str) and ("total" in c.lower() or "26.1" in str(c))]:
                    chosen = cand
                    break
                if chosen is None and len(candidate_cols) > 0:
                    chosen = candidate_cols[0]

                # if chosen column refers to matr or leads, try to map
                val = int(row[chosen].sum()) if not row[chosen].isnull().all() else None

                # heuristically assign to leads or matriculas based on sheet name
                name_low = name.lower()
                if "mat" in name_low or "matric" in name_low or (isinstance(chosen, str) and "mat" in chosen.lower()):
                    res[u]["mat_excel"] = val
                    res[u]["sheet_mat"] = f"{name} -> {chosen}"
                else:
                    # try leads assignment
                    res[u]["leads_excel"] = val
                    res[u]["sheet_leads"] = f"{name} -> {chosen}"

    return res


def main():
    print("Lendo CSV...", CSV_PATH)
    if not CSV_PATH.exists():
        print("Arquivo CSV não encontrado:", CSV_PATH)
        return
    df_csv = read_csv_attempt(CSV_PATH)

    print("Contando no CSV...")
    csv_res = csv_counts(df_csv, UNITS)

    print("Lendo Excel (tentativa de extrair por-unidade)...", EXCEL_PATH)
    excel_res = excel_extract_units(EXCEL_PATH, UNITS) if EXCEL_PATH.exists() else {u: {} for u in UNITS}

    print("\nResultado comparativo (CSV vs Excel)\n")
    for u in UNITS:
        c = csv_res.get(u, {})
        e = excel_res.get(u, {})
        print(f"Unidade: {u}")
        print(f" - Leads (CSV)   : {c.get('leads_total_csv')} | Leads ciclo 26.1 (CSV): {c.get('leads_26_csv')}")
        print(f" - Matriculas (CSV): {c.get('mat_total_csv')} | Matriculas ciclo 26.1 (CSV): {c.get('mat_26_csv')}")
        print(f" - Leads (Excel) : {e.get('leads_excel')} ({e.get('sheet_leads')})")
        print(f" - Matriculas (Excel): {e.get('mat_excel')} ({e.get('sheet_mat')})")
        print("")


if __name__ == "__main__":
    main()
