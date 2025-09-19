import io
import sys
import math
import json
import textwrap
from datetime import datetime
from typing import List, Optional

import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px

# ------------- CONFIG -------------
st.set_page_config(page_title="Relatórios por Secretaria", layout="wide")
pd.options.display.float_format = lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ------------- HELPERS -------------
def brl(x) -> str:
    try:
        return f"R$ {float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)

def parse_dates(series: pd.Series) -> pd.Series:
    # Tenta converter datas em vários formatos, incluindo dd/mm/aaaa
    return pd.to_datetime(series, errors="coerce", dayfirst=True, infer_datetime_format=True)

def unify_amounts(df: pd.DataFrame, amount_col: str, type_col: Optional[str], credit_tokens: List[str], debit_tokens: List[str]) -> pd.Series:
    """
    Retorna uma série com valores 'assinados':
      - Se não houver coluna de tipo, retorna o próprio valor (assumindo já estar positivo/negativo).
      - Se houver, mapeia tokens de crédito como +abs(valor) e débito como -abs(valor).
    """
    vals = pd.to_numeric(df[amount_col], errors="coerce")
    if type_col is None:
        return vals

    t = df[type_col].astype(str).str.strip().str.upper()
    credit_set = {s.strip().upper() for s in credit_tokens if s.strip()}
    debit_set  = {s.strip().upper() for s in debit_tokens  if s.strip()}

    signed = vals.copy()
    # Por padrão, mantém o sinal existente se não bater em nenhum token
    # Aplica sinais quando bater
    is_credit = t.isin(credit_set)
    is_debit  = t.isin(debit_set)

    signed[is_credit] = vals[is_credit].abs()
    signed[is_debit]  = -vals[is_debit].abs()
    return signed

def to_excel_bytes(df_dict: dict) -> bytes:
    """
    Recebe dict {nome_aba: dataframe} e devolve bytes de um .xlsx formatado básico.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            for idx, col in enumerate(df.columns):
                # largura automática aproximada
                max_len = max(12, df[col].astype(str).map(len).max() if not df.empty else 12)
                ws.set_column(idx, idx, min(max_len + 2, 50))
    return output.getvalue()

# ------------- SIDEBAR / INPUT -------------
st.sidebar.header("Importação de Dados")

default_path = "/mnt/data/Saldos BB 2025 - 11-09.xlsx"  # disponível se você subir o arquivo no chat
use_default = False

uploaded_file = st.sidebar.file_uploader("Envie sua planilha (Excel .xlsx)", type=["xlsx"])
if uploaded_file is None:
    st.sidebar.markdown("Ou use o arquivo padrão detectado no ambiente:")
    use_default = st.sidebar.toggle("Usar arquivo padrão do ambiente", value=True)
    if use_default:
        try:
            with open(default_path, "rb") as f:
                uploaded_file = io.BytesIO(f.read())
            st.sidebar.success("Arquivo padrão carregado.")
        except Exception:
            st.sidebar.warning("Arquivo padrão não encontrado. Envie um .xlsx acima.")
else:
    st.sidebar.success("Arquivo enviado.")

if uploaded_file is None:
    st.info("Envie um arquivo .xlsx (ou habilite o padrão na barra lateral).")
    st.stop()

# Descoberta de abas
try:
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names
except Exception as e:
    st.error(f"Falha ao abrir o Excel: {e}")
    st.stop()

st.sidebar.subheader("Seleção de Planilha")
sheet_name = st.sidebar.selectbox("Aba a importar", options=sheets, index=0)

try:
    df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name)
except Exception as e:
    st.error(f"Falha ao ler a aba '{sheet_name}': {e}")
    st.stop()

if df_raw.empty:
    st.warning("A aba selecionada está vazia.")
    st.stop()

# Normaliza nomes de colunas visualmente (sem alterar o original)
columns = list(df_raw.columns)

st.sidebar.subheader("Mapeamento de Colunas")
col_secretaria = st.sidebar.selectbox("Coluna de Secretaria (obrigatório)", options=columns)
col_valor      = st.sidebar.selectbox("Coluna de Valor (obrigatório)", options=columns)
col_data_opt   = st.sidebar.selectbox("Coluna de Data (opcional)", options=["<Nenhuma>"] + columns, index=0)
col_tipo_opt   = st.sidebar.selectbox("Coluna de Tipo Crédito/Débito (opcional)", options=["<Nenhuma>"] + columns, index=0)

has_date = col_data_opt != "<Nenhuma>"
has_type = col_tipo_opt != "<Nenhuma>"

credit_tokens = []
debit_tokens = []
if has_type:
    with st.sidebar.expander("Mapeamento de Tipo (opcional)", expanded=False):
        st.markdown(
            "Informe como sua planilha marca **Crédito** e **Débito**.\n\n"
            "- Ex.: crédito = `C, CR, CREDITO`; débito = `D, DB, DEBITO`"
        )
        credit_tokens = [s for s in st.text_input("Tokens de Crédito (separados por vírgula)", value="C, CR, CREDITO").split(",")]
        debit_tokens  = [s for s in st.text_input("Tokens de Débito (separados por vírgula)", value="D, DB, DEBITO").split(",")]

# ------------- PREPARO DOS DADOS -------------
df = df_raw.copy()

# Data
if has_date:
    df["_date"] = parse_dates(df[col_data_opt])
else:
    df["_date"] = pd.NaT

# Valor assinado
signed_amount = unify_amounts(df, amount_col=col_valor, type_col=(col_tipo_opt if has_type else None),
                              credit_tokens=credit_tokens, debit_tokens=debit_tokens)
df["_amount"] = pd.to_numeric(signed_amount, errors="coerce")

# Secretaria
df["_dept"] = df[col_secretaria].astype(str).str.strip()

# Limpa registros sem valor numérico ou sem secretaria
df = df[~df["_dept"].isna() & ~df["_dept"].eq("")]
df = df[~df["_amount"].isna()]

# ------------- FILTROS -------------
st.sidebar.header("Filtros")
dept_list = sorted(df["_dept"].unique().tolist())
selected_depts = st.sidebar.multiselect("Secretarias", options=dept_list, default=dept_list)

if has_date and df["_date"].notna().any():
    min_date = pd.to_datetime(df["_date"].min()).date()
    max_date = pd.to_datetime(df["_date"].max()).date()
    start_date, end_date = st.sidebar.date_input(
        "Período (Data)",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )
else:
    start_date, end_date = None, None

df_f = df.copy()
if selected_depts:
    df_f = df_f[df_f["_dept"].isin(selected_depts)]

if has_date and start_date and end_date:
    mask = (df_f["_date"].dt.date >= start_date) & (df_f["_date"].dt.date <= end_date)
    df_f = df_f[mask]

if df_f.empty:
    st.warning("Nenhum dado após aplicar os filtros.")
    st.stop()

# ------------- MÉTRICAS E RELATÓRIOS -------------
st.title("Relatórios e Gráficos por Secretaria")

left, right = st.columns([2, 1], gap="large")

with left:
    st.subheader("Totais por Secretaria")
    by_dept = df_f.groupby("_dept", as_index=False)["_amount"].sum().rename(columns={"_dept": "Secretaria", "_amount": "Valor"})
    by_dept_sorted = by_dept.sort_values("Valor", ascending=False).reset_index(drop=True)

    fig = px.bar(by_dept_sorted, x="Secretaria", y="Valor", text="Valor")
    fig.update_traces(texttemplate="%{text:.2f}", textposition="outside")
    fig.update_layout(yaxis_title="Valor (R$)", xaxis_title="Secretaria", uniformtext_minsize=8, uniformtext_mode="hide", bargap=0.3)
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(
        by_dept_sorted.assign(Valor=by_dept_sorted["Valor"].map(brl)),
        use_container_width=True
    )

    total_geral = by_dept_sorted["Valor"].sum()
    st.markdown(f"**Total Geral (filtros aplicados): {brl(total_geral)}**")

with right:
    st.subheader("Resumo Rápido")
    n_regs = len(df_f)
    n_depts = df_f["_dept"].nunique()
    if has_date and df_f["_date"].notna().any():
        period_info = f"Período: **{df_f['_date'].min().date().strftime('%d/%m/%Y')}** a **{df_f['_date'].max().date().strftime('%d/%m/%Y')}**"
    else:
        period_info = "Período: **não informado**"

    st.metric("Registros", n_regs)
    st.metric("Secretarias distintas", n_depts)
    st.metric("Total geral", brl(df_f["_amount"].sum()))
    st.caption(period_info)

# Série temporal por mês (se houver data)
if has_date and df_f["_date"].notna().any():
    st.subheader("Evolução Mensal")
    df_f["_year_month"] = df_f["_date"].dt.to_period("M").dt.to_timestamp()
    monthly = df_f.groupby(["_year_month", "_dept"], as_index=False)["_amount"].sum()
    # Gráfico empilhado por secretaria
    pivot = monthly.pivot(index="_year_month", columns="_dept", values="_amount").fillna(0.0)
    pivot = pivot.sort_index()
    fig2 = px.area(pivot, x=pivot.index, y=pivot.columns)
    fig2.update_layout(xaxis_title="Mês", yaxis_title="Valor (R$)")
    st.plotly_chart(fig2, use_container_width=True)

# Tabela detalhada (com colunas originais úteis)
st.subheader("Relatório Detalhado (com filtros)")
show_cols = []
# Mostra colunas relevantes primeiro
preferred = [col_secretaria, col_valor]
if has_date:
    preferred.append(col_data_opt)
if has_type:
    preferred.append(col_tipo_opt)

# Evita duplicatas preservando ordem
for c in preferred:
    if c in df_f.columns and c not in show_cols:
        show_cols.append(c)

# Acrescenta uma coluna Valor (assinado) formatado e Data formatada
df_report = df_f.copy()
df_report["Valor (assinado)"] = df_report["_amount"].map(brl)
if has_date:
    df_report["Data (dd/mm/aaaa)"] = df_report["_date"].dt.strftime("%d/%m/%Y")

# Completa com demais colunas originais (sem as técnicas iniciadas com "_")
for c in df_f.columns:
    if not c.startswith("_") and c not in show_cols:
        show_cols.append(c)

# Adiciona colunas calculadas no final
extra_cols = []
if "Valor (assinado)" not in show_cols:
    extra_cols.append("Valor (assinado)")
if has_date and "Data (dd/mm/aaaa)" not in show_cols:
    extra_cols.append("Data (dd/mm/aaaa)")

final_cols = [c for c in show_cols if c in df_report.columns] + extra_cols
st.dataframe(df_report[final_cols], use_container_width=True)

# ------------- DOWNLOADS -------------
st.subheader("Exportar Relatórios")

export_summary = by_dept_sorted.copy()
export_summary["Valor"] = export_summary["Valor"].round(2)

export_detail = df_report.copy()
# Colunas técnicas saem na exportação
export_detail = export_detail[[c for c in export_detail.columns if not c.startswith("_")]]

col_a, col_b = st.columns(2)
with col_a:
    csv_bytes = export_summary.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="Baixar Resumo por Secretaria (CSV)",
        data=csv_bytes,
        file_name="resumo_por_secretaria.csv",
        mime="text/csv"
    )
with col_b:
    # Excel com abas
    excel_bytes = to_excel_bytes({
        "Resumo_por_Secretaria": export_summary,
        "Detalhado": export_detail
    })
    st.download_button(
        label="Baixar Relatório Completo (Excel)",
        data=excel_bytes,
        file_name="relatorio_por_secretaria.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.caption("Dica: mapeie as colunas corretamente na barra lateral. Se sua planilha tiver 'C/D' ou 'Crédito/Débito', informe os tokens para assinar os valores.")
