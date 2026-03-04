"""
Análise de Negócios Perdidos - Red Balloon
Espelho da lógica de qualificação, focado em leads na etapa de Negócio Perdido.
"""

from pathlib import Path
from datetime import datetime
import pandas as pd


print("📊 Iniciando análise de Negócios Perdidos...")
BASE_DIR = Path(__file__).parent.parent

candidate_paths = [
    Path(r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\hubspot_leads_ATUAL.csv"),
    BASE_DIR / "data" / "hubspot_leads_redballoon.csv",
    Path(r"C:\Users\a483650\Projetos\_DADOS_CENTRALIZADOS\hubspot\historico\hubspot_leads_ATUAL.csv"),
    Path(r"C:\Users\a483650\Projetos\_ARQUIVO\projeto_helio_teste\Data\hubspot_leads_atual.csv"),
]

found_file = None
for path in candidate_paths:
    if path.exists():
        found_file = path
        break

if found_file is None:
    raise FileNotFoundError(
        "Nenhum arquivo de leads encontrado. Verifique caminhos: " + ", ".join([str(p) for p in candidate_paths])
    )

print(f"📥 Usando base de dados: {found_file}")
df = pd.read_csv(found_file)

required_columns = [
    "Etapa do negócio",
    "Data de criação",
    "Data da última atividade",
    "Data de fechamento",
    "Número de atividades de vendas",
    "Fonte original do tráfego",
    "Unidade Desejada",
]

missing = [column for column in required_columns if column not in df.columns]
if missing:
    raise ValueError(f"Colunas obrigatórias ausentes na base: {missing}")

# Normalização de datas
df["Data de criação"] = pd.to_datetime(df["Data de criação"], errors="coerce")
df["Data da última atividade"] = pd.to_datetime(df["Data da última atividade"], errors="coerce")
df["Data de fechamento"] = pd.to_datetime(df["Data de fechamento"], errors="coerce")

# Garantir tipo numérico para atividades
df["Número de atividades de vendas"] = pd.to_numeric(df["Número de atividades de vendas"], errors="coerce").fillna(0)


def definir_ciclo(data: pd.Timestamp):
    if pd.isna(data):
        return None
    ano = data.year
    mes = data.month

    if mes >= 10:
        return f"{str(ano + 1)[-2:]}.1"
    if mes <= 3:
        return f"{str(ano)[-2:]}.1"
    return None


def classificar_atividades(atividades: float):
    if atividades == 0:
        return "0 atividades"
    if atividades <= 2:
        return "1-2 atividades"
    if atividades <= 5:
        return "3-5 atividades"
    if atividades <= 10:
        return "6-10 atividades"
    return "11+ atividades"


def classificar_idade(dias: float):
    if pd.isna(dias):
        return "Sem data"
    if dias <= 30:
        return "0-30 dias"
    if dias <= 60:
        return "31-60 dias"
    if dias <= 90:
        return "61-90 dias"
    if dias <= 180:
        return "91-180 dias"
    return "180+ dias"


def classificar_recencia(dias: float):
    if pd.isna(dias):
        return "Nunca contatado"
    if dias <= 7:
        return "Ativos (0-7 dias)"
    if dias <= 30:
        return "Recentes (8-30 dias)"
    if dias <= 60:
        return "Mornos (31-60 dias)"
    if dias <= 90:
        return "Frios (61-90 dias)"
    return "Muito Frios (>90 dias)"


# Classificações base
df["ciclo"] = df["Data de criação"].apply(definir_ciclo)
df["canal"] = df["Fonte original do tráfego"].fillna("Não informado")
df["unidade"] = df["Unidade Desejada"].fillna("Não informada")

# Filtro de Negócio Perdido
mask_perdido = df["Etapa do negócio"].astype(str).str.upper().str.contains("PERDIDO", na=False)
perdidos = df[mask_perdido].copy()

agora = datetime.now()
perdidos["Dias_Desde_Criacao"] = (agora - perdidos["Data de criação"]).dt.days
perdidos["Dias_Sem_Contato"] = (agora - perdidos["Data da última atividade"]).dt.days
perdidos["Dias_Sem_Contato"] = perdidos["Dias_Sem_Contato"].fillna(999).astype(int)
perdidos["Dias_Para_Perder"] = (perdidos["Data de fechamento"] - perdidos["Data de criação"]).dt.days

perdidos["Faixa_Atividades"] = perdidos["Número de atividades de vendas"].apply(classificar_atividades)
perdidos["Faixa_Idade"] = perdidos["Dias_Desde_Criacao"].apply(classificar_idade)
perdidos["Faixa_Recencia"] = perdidos["Dias_Sem_Contato"].apply(classificar_recencia)

# Colunas opcionais úteis
optional_columns = [
    "Record ID",
    "Nome do negócio",
    "Nome da empresa",
    "Contato",
    "Telefone",
    "E-mail",
    "Proprietário do negócio",
    "Motivo da perda",
    "Descrição do motivo da perda",
]

base_columns = [
    "Etapa do negócio",
    "Data de criação",
    "Data de fechamento",
    "ciclo",
    "unidade",
    "canal",
    "Número de atividades de vendas",
    "Data da última atividade",
    "Dias_Desde_Criacao",
    "Dias_Sem_Contato",
    "Dias_Para_Perder",
    "Faixa_Atividades",
    "Faixa_Idade",
    "Faixa_Recencia",
]

existing_optional = [column for column in optional_columns if column in perdidos.columns]
selected_columns = existing_optional + base_columns

perdidos_lista = perdidos[selected_columns].copy()
perdidos_lista = perdidos_lista.sort_values(
    ["Data de fechamento", "Número de atividades de vendas", "Dias_Sem_Contato"],
    ascending=[False, False, True],
)

# Métricas gerais
total_perdidos = len(perdidos_lista)
sem_atividade = int((perdidos_lista["Número de atividades de vendas"] == 0).sum()) if total_perdidos > 0 else 0
com_atividade = int((perdidos_lista["Número de atividades de vendas"] > 0).sum()) if total_perdidos > 0 else 0
nunca_contatados = int((perdidos_lista["Dias_Sem_Contato"] >= 999).sum()) if total_perdidos > 0 else 0

resumo_geral = pd.DataFrame(
    [
        {
            "Métrica": "Total Negócios Perdidos",
            "Valor": total_perdidos,
        },
        {
            "Métrica": "Perdidos com atividade de vendas (>0)",
            "Valor": com_atividade,
        },
        {
            "Métrica": "Perdidos sem atividade (0)",
            "Valor": sem_atividade,
        },
        {
            "Métrica": "% com atividade",
            "Valor": round((com_atividade / total_perdidos * 100), 1) if total_perdidos > 0 else 0,
        },
        {
            "Métrica": "Nunca contatados",
            "Valor": nunca_contatados,
        },
        {
            "Métrica": "% nunca contatados",
            "Valor": round((nunca_contatados / total_perdidos * 100), 1) if total_perdidos > 0 else 0,
        },
        {
            "Métrica": "Tempo médio até perder (dias)",
            "Valor": round(perdidos_lista["Dias_Para_Perder"].dropna().mean(), 1)
            if total_perdidos > 0 and perdidos_lista["Dias_Para_Perder"].notna().any()
            else None,
        },
    ]
)

por_ciclo = (
    perdidos_lista.groupby("ciclo", dropna=False)
    .agg(
        Total=("Etapa do negócio", "size"),
        Com_Atividade=("Número de atividades de vendas", lambda series: (series > 0).sum()),
        Media_Atividades=("Número de atividades de vendas", "mean"),
        Media_Dias_Para_Perder=("Dias_Para_Perder", "mean"),
    )
    .reset_index()
    .sort_values("ciclo", ascending=False)
)

por_canal = (
    perdidos_lista.groupby("canal", dropna=False)
    .agg(
        Total=("Etapa do negócio", "size"),
        Com_Atividade=("Número de atividades de vendas", lambda series: (series > 0).sum()),
        Media_Dias_Para_Perder=("Dias_Para_Perder", "mean"),
    )
    .reset_index()
    .sort_values("Total", ascending=False)
)

por_unidade = (
    perdidos_lista.groupby("unidade", dropna=False)
    .agg(
        Total=("Etapa do negócio", "size"),
        Com_Atividade=("Número de atividades de vendas", lambda series: (series > 0).sum()),
        Media_Dias_Para_Perder=("Dias_Para_Perder", "mean"),
    )
    .reset_index()
    .sort_values("Total", ascending=False)
)

analise_recencia = (
    perdidos_lista.groupby("Faixa_Recencia", dropna=False)
    .size()
    .reset_index(name="Quantidade")
    .sort_values("Quantidade", ascending=False)
)
if total_perdidos > 0:
    analise_recencia["Percentual"] = (analise_recencia["Quantidade"] / total_perdidos * 100).round(1)
else:
    analise_recencia["Percentual"] = 0

analise_faixa_idade = (
    perdidos_lista.groupby("Faixa_Idade", dropna=False)
    .size()
    .reset_index(name="Quantidade")
    .sort_values("Quantidade", ascending=False)
)
if total_perdidos > 0:
    analise_faixa_idade["Percentual"] = (analise_faixa_idade["Quantidade"] / total_perdidos * 100).round(1)
else:
    analise_faixa_idade["Percentual"] = 0

analise_atividades = (
    perdidos_lista.groupby("Faixa_Atividades", dropna=False)
    .size()
    .reset_index(name="Quantidade")
    .sort_values("Quantidade", ascending=False)
)
if total_perdidos > 0:
    analise_atividades["Percentual"] = (analise_atividades["Quantidade"] / total_perdidos * 100).round(1)
else:
    analise_atividades["Percentual"] = 0

# Exportar para Excel
output_path = BASE_DIR / "outputs" / "Analise_Negocios_Perdidos.xlsx"
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    resumo_geral.to_excel(writer, sheet_name="Resumo_Geral", index=False)
    por_ciclo.to_excel(writer, sheet_name="Perdidos_Por_Ciclo", index=False)
    por_canal.to_excel(writer, sheet_name="Perdidos_Por_Canal", index=False)
    por_unidade.to_excel(writer, sheet_name="Perdidos_Por_Unidade", index=False)
    analise_recencia.to_excel(writer, sheet_name="Recencia_Sem_Contato", index=False)
    analise_faixa_idade.to_excel(writer, sheet_name="Faixa_Idade", index=False)
    analise_atividades.to_excel(writer, sheet_name="Faixa_Atividades", index=False)
    perdidos_lista.to_excel(writer, sheet_name="Lista_Negocios_Perdidos", index=False)

print("\n✅ Análise concluída com sucesso!")
print(f"📁 Arquivo salvo em: {output_path}")
print("\n📌 Resumo:")
print(f"   Total de Negócios Perdidos: {total_perdidos:,}")
print(f"   Perdidos com atividade (>0): {com_atividade:,}")
print(f"   Perdidos sem atividade: {sem_atividade:,}")
print(f"   Nunca contatados: {nunca_contatados:,}")
