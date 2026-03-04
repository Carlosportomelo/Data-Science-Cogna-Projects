import pandas as pd
from pathlib import Path
import sys
try:
    import xlrd
except Exception:
    xlrd = None

BASE_DIR  = Path(r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\receita")
F_PAG     = BASE_DIR / "Relatorio_DadosPagamentos.xlsx"
F_ABERTO  = BASE_DIR / "parcelas_em_aberto_27_02.xls"
F_ESTORNO = BASE_DIR / "estornos_27_02.xls"
OUT_DIR   = Path(r"C:\Users\a483650\Projetos\07_Redballoon_RECEITA_RVO\Outputs")
OUTPUT    = OUT_DIR / "desconto_por_unidade.xlsx"

def ler(path):
    p = Path(path)
    if not p.exists():
        print(f"[ERRO] Arquivo não encontrado: {p}"); sys.exit(1)
    suffix = p.suffix.lower()
    # .xls -> usar xlrd diretamente para evitar conflitos com pandas
    if suffix == ".xls":
        # tentativa com xlrd
        if xlrd is not None:
            try:
                wb = xlrd.open_workbook(str(p))
                sh = wb.sheet_by_index(0)
                headers = [str(c).strip() for c in sh.row_values(0)]
                data = [sh.row_values(r) for r in range(1, sh.nrows)]
                df = pd.DataFrame(data, columns=headers)
                df.columns = df.columns.str.strip()
                return df
            except Exception:
                # fallthrough to try HTML parsing / pandas
                pass
        # Se não foi possível com xlrd, tentar interpretar como HTML (alguns .xls são HTML)
        try:
            dfs = pd.read_html(str(p), header=0)
            if len(dfs) >= 1:
                df = dfs[0]
            else:
                raise ValueError("Nenhuma tabela encontrada em HTML")
        except Exception as e:
            print(f"[ERRO] Não foi possível ler .xls (xlrd e read_html falharam): {e}")
            print("Se for um arquivo Excel real, instale 'xlrd==1.2.0' ou converta para .xlsx")
            sys.exit(1)
    else:
        try:
            df = pd.read_excel(p, engine="openpyxl")
        except Exception:
            try:
                df = pd.read_excel(p)
            except Exception as e:
                print(f"[ERRO] Não foi possível ler o arquivo {p}: {e}")
                sys.exit(1)
    df.columns = df.columns.str.strip()
    return df

df_pag    = ler(F_PAG)
df_aberto = ler(F_ABERTO)
df_est    = ler(F_ESTORNO)

# Pagamentos: Valor Original vs Valor Pago
df_pag["Valor_Original"] = pd.to_numeric(df_pag["Valor Original"], errors="coerce")
df_pag["Valor_Final"]    = pd.to_numeric(df_pag["Valor Pago"],     errors="coerce")

# Em Aberto: Valor Original vs Valor a Pagar
df_aberto["Valor_Original"] = pd.to_numeric(df_aberto["Valor Original"], errors="coerce")
df_aberto["Valor_Final"]    = pd.to_numeric(df_aberto["Valor a Pagar"],  errors="coerce")

# Estornos: Valor Original vs Valor Estornado
val_col = "Valor Estornado" if "Valor Estornado" in df_est.columns else "Valor a Pagar"
df_est["Valor_Original"] = pd.to_numeric(df_est["Valor Original"], errors="coerce")
df_est["Valor_Final"]    = pd.to_numeric(df_est[val_col],          errors="coerce")

# Empilha as 3 bases mantendo identificadores individuais (matrícula/registro)
def detect_id_column(dfs):
    candidates = ["Matrícula", "Matricula", "matricula", "Matricula ID", "Contrato", "Contrato Nº", "ContratoNum", "Contrato Nº", "ID", "Registro", "RegistroID"]
    for c in candidates:
        for df in dfs:
            if c in df.columns:
                return c
    return None

id_col = detect_id_column([df_pag, df_aberto, df_est])

parts = []
for df in (df_pag, df_aberto, df_est):
    cols = [c for c in ["Unidade", id_col, "Valor_Original", "Valor_Final"] if c in df.columns]
    parts.append(df[cols].copy())

rol = pd.concat(parts, ignore_index=True)

# Normaliza nomes e remove linhas sem valores
if id_col is None:
    rol = rol.reset_index().rename(columns={"index": "RegistroID"})
    id_col = "RegistroID"

rol = rol.dropna(subset=["Unidade", "Valor_Original"]).copy()
rol["Valor_Final"] = rol["Valor_Final"].fillna(0)

# Calcula desconto por registro
rol["Desconto (R$)"] = rol["Valor_Original"] - rol["Valor_Final"]
rol["Desconto (%)"] = rol.apply(lambda r: (r["Desconto (R$)"] / r["Valor_Original"]) if r["Valor_Original"] not in (0, None, float('nan')) else 0, axis=1)

# Reordena colunas de saída
out_cols = ["Unidade", id_col, "Valor_Original", "Valor_Final", "Desconto (R$)", "Desconto (%)"]
resumo = rol[[c for c in out_cols if c in rol.columns]].copy()

# Formatação: renomear colunas para português padrão de saída
resumo = resumo.rename(columns={
    id_col: "Matricula",
    "Valor_Original": "Mensalidade (R$)",
    "Valor_Final": "Valor Final (R$)",
})

# Normalizar matrícula para reduzir variações (espacos, cases, floats)
resumo["Matricula"] = resumo["Matricula"].astype(str).str.strip()
resumo.loc[resumo["Matricula"].str.lower().isin(["nan", "none", "nan.0", "nan.00", "nan.000"]), "Matricula"] = ""
resumo = resumo[resumo["Matricula"] != ""].copy()
resumo = resumo.sort_values(["Unidade", "Matricula"]).reset_index(drop=True)

# Separar registros com desconto entre 90% e 100% (não serão considerados nas agregações)
high_mask = resumo["Desconto (%)"] >= 0.9
high_discount = resumo[high_mask].copy()
resumo = resumo[~high_mask].copy()

# Agrupar por Unidade + combinação de Mensalidade e Valor Final para mostrar volumes
agg = resumo.groupby(["Unidade", "Mensalidade (R$)", "Valor Final (R$)"]).agg(
    Count=("Matricula", pd.Series.nunique),
    Desconto_R=("Desconto (R$)", "mean"),
    Desconto_pct=("Desconto (%)", "mean"),
).reset_index()

# Percentual de cada grupo dentro da unidade
agg["Total_Por_Unidade"] = agg.groupby("Unidade")["Count"].transform("sum")
agg["Pct_da_Unidade"] = agg["Count"] / agg["Total_Por_Unidade"]

# Salva agregação
agg_xlsx = OUT_DIR / "desconto_agrupado_por_unidade.xlsx"
agg_csv = OUT_DIR / "desconto_agrupado_por_unidade.csv"
agg.to_excel(agg_xlsx, index=False)
agg.to_csv(agg_csv, index=False)

# Impressão compacta por unidade
current_unit = None
for _, row in agg.sort_values(["Unidade", "Count"], ascending=[True, False]).iterrows():
    u = row["Unidade"]
    c = int(row["Count"])
    mens = row["Mensalidade (R$)"]
    paid = row["Valor Final (R$)"]
    dr = row["Desconto_R"]
    dp = row["Desconto_pct"]
    pct_unit = row["Pct_da_Unidade"]
    print(f"{u}: {c} matrícula(s) com mensalidade {mens:.2f}, pago {paid:.2f} → desconto {dr:.2f} ({dp:.2%}) — {pct_unit:.1%} do total da unidade")

print(f"✅ Agrupado salvo em: {agg_xlsx}")
print(f"✅ Agrupado CSV salvo em: {agg_csv}")

# 1) Resumo por unidade (médias) — equivalente à versão anterior
resumo_media = resumo.groupby("Unidade").agg(
    Mensalidade_Media=("Mensalidade (R$)", "mean"),
    Valor_Final_Medio=("Valor Final (R$)", "mean"),
    Qtde_Contratos=("Matricula", pd.Series.nunique),
).reset_index()

resumo_media["Desconto (R$)"] = resumo_media["Mensalidade_Media"] - resumo_media["Valor_Final_Medio"]
resumo_media["Desconto (%)"] = resumo_media["Desconto (R$)"] / resumo_media["Mensalidade_Media"]

# Salva resumo médio por unidade
OUT_DIR.mkdir(parents=True, exist_ok=True)
media_xlsx = OUT_DIR / "desconto_por_unidade_media.xlsx"
media_csv = OUT_DIR / "desconto_por_unidade_media.csv"
try:
    resumo_media.to_excel(media_xlsx, index=False)
    print(f"✅ Resumo médio salvo em: {media_xlsx}")
except PermissionError:
    print(f"[AVISO] Não foi possível sobrescrever {media_xlsx} (arquivo pode estar aberto). CSV será salvo.")
resumo_media.to_csv(media_csv, index=False)
print(f"✅ Resumo médio CSV salvo em: {media_csv}")

# 2) Mantém o relatório detalhado por matrícula (resumo já produzido acima)
for _, r in resumo.iterrows():
    unidade = r["Unidade"]
    mensal = r["Mensalidade (R$)"]
    final = r["Valor Final (R$)"]
    desc_r = r["Desconto (R$)"]
    desc_pct = r["Desconto (%)"]
    print(f"Unidade {unidade}, Mensalidade é {mensal:.2f}, Aplica desconto {desc_r:.2f} ({desc_pct:.2%}), Valor final {final:.2f}")

# Também salva o detalhado (já salvo anteriormente como desconto_por_unidade.xlsx/csv)
csv_out = OUT_DIR / "desconto_por_unidade.csv"
print(f"✅ Detalhado salvo em: {OUTPUT} e {csv_out}")

# 3) Resumo por matrícula (consolidado) — MÉDIA por matrícula (conforme solicitado)
por_matricula = resumo.groupby("Matricula").agg(
    Unidade=("Unidade", lambda s: s.mode().iat[0] if not s.mode().empty else s.iloc[0]),
    Qtde_Registros=("Mensalidade (R$)", "count"),
    Mensalidade_Media=("Mensalidade (R$)", "mean"),
    Valor_Final_Medio=("Valor Final (R$)", "mean"),
    Desconto_Medio=("Desconto (R$)", "mean"),
    Desconto_pct_Medio=("Desconto (%)", "mean"),
).reset_index()

por_matricula = por_matricula.sort_values("Qtde_Registros", ascending=False)
mat_xlsx = OUT_DIR / "desconto_por_matricula_media.xlsx"
mat_csv = OUT_DIR / "desconto_por_matricula_media.csv"
por_matricula.to_excel(mat_xlsx, index=False)
por_matricula.to_csv(mat_csv, index=False)

# Salvar consolidado por matrícula e aba separada para descontos 90-100%
with pd.ExcelWriter(mat_xlsx) as w:
    por_matricula.to_excel(w, sheet_name="Consolidado", index=False)
    if not high_discount.empty:
        high_discount.to_excel(w, sheet_name="Desconto_90_100", index=False)

por_matricula.to_csv(mat_csv, index=False)
high_csv = OUT_DIR / "desconto_90_100.csv"
if not high_discount.empty:
    high_discount.to_csv(high_csv, index=False)

print(f"✅ Consolidado por matrícula salvo em: {mat_xlsx}")
print(f"✅ CSV por matrícula salvo em: {mat_csv}")
if not high_discount.empty:
    print(f"✅ Registros com 90-100% de desconto salvos em aba 'Desconto_90_100' e em: {high_csv}")
print(f"🔢 Matrículas únicas no consolidado: {por_matricula['Matricula'].nunique()}")