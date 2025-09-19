import io
from datetime import datetime
import re

import pandas as pd
import plotly.express as px
import streamlit as st

# =============== CONFIG ===============
st.set_page_config(page_title="Saldos BB 2025 - Secretaria e Relatório", layout="wide")
pd.options.display.float_format = lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# =============== HELPERS ===============
def brl(x) -> str:
    try:
        return f"R$ {float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)

def load_all_sheets(xlsx_bytes_or_path) -> pd.DataFrame:
    """
    Lê todas as abas (cada aba = um dia no formato dd-mm-aaaa) e retorna DF consolidado:
    ['Conta','Nome da Conta','Secretaria','Banco','Tipo de Recurso','Saldo Bancario','Date']
    """
    xls = pd.ExcelFile(xlsx_bytes_or_path)
    frames = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df.columns = [c.strip() for c in df.columns]

        expected = {"Conta", "Nome da Conta", "Secretaria", "Banco", "Tipo de Recurso", "Saldo Bancario"}
        missing = expected.difference(df.columns)
        if missing:
            raise ValueError(f"Aba '{sheet}' sem colunas esperadas: {missing}")

        # Data a partir do nome da aba
        d = pd.to_datetime(sheet, format="%d-%m-%Y", dayfirst=True, errors="coerce")
        df["Date"] = d

        frames.append(df[["Conta","Nome da Conta","Secretaria","Banco","Tipo de Recurso","Saldo Bancario","Date"]])

    out = pd.concat(frames, ignore_index=True)
    # Tipagens e limpeza
    out["Saldo Bancario"] = pd.to_numeric(out["Saldo Bancario"], errors="coerce")
    out["Secretaria"] = out["Secretaria"].astype(str).str.strip()
    out["Conta"] = out["Conta"].astype(str).str.strip()
    out["Nome da Conta"] = out["Nome da Conta"].astype(str).str.strip()
    out["Banco"] = out["Banco"].astype(str).str.strip()
    out["Tipo de Recurso"] = out["Tipo de Recurso"].astype(str).str.strip()
    out = out.dropna(subset=["Saldo Bancario"])
    return out

def only_digits(text: str) -> str:
    """Mantém apenas dígitos."""
    return re.sub(r"\D+", "", text or "")

def conta_prefix(conta: str) -> str:
    """
    Retorna a parte antes do hífen, somente dígitos.
    Ex.: '12345-0' -> '12345'
    """
    if conta is None:
        return ""
    # pega tudo antes do primeiro '-'
    head = str(conta).split("-")[0]
    return only_digits(head)

# =============== INPUT ===============
st.sidebar.header("Arquivo")
default_path = "/mnt/data/09 - Saldos BB 2025.xlsx"

uploaded = st.sidebar.file_uploader("Excel (.xlsx) com abas por data (dd-mm-aaaa)", type=["xlsx"])
use_default = st.sidebar.toggle("Usar arquivo padrão do ambiente", value=(uploaded is None))

if uploaded:
    source = uploaded
elif use_default:
    try:
        with open(default_path, "rb") as f:
            source = io.BytesIO(f.read())
        st.sidebar.success("Arquivo padrão carregado.")
    except Exception:
        st.sidebar.error("Arquivo padrão não encontrado. Envie o .xlsx.")
        st.stop()
else:
    st.info("Envie a planilha ou ative o arquivo padrão.")
    st.stop()

# =============== LOAD ===============
try:
    df = load_all_sheets(source)
except Exception as e:
    st.error(f"Falha ao carregar planilha: {e}")
    st.stop()

if df.empty:
    st.warning("Sem dados.")
    st.stop()

# =============== FILTROS BÁSICOS ===============
st.title("Saldos BB 2025 — Gráfico por Secretaria + Relatório")

# Período
if df["Date"].notna().any():
    min_d = df["Date"].min().date()
    max_d = df["Date"].max().date()
    d_ini, d_fim = st.sidebar.date_input("Período", value=(min_d, max_d), min_value=min_d, max_value=max_d)
else:
    d_ini, d_fim = None, None

# Filtro por Secretaria (multiselect)
sec_opts = sorted(df["Secretaria"].dropna().unique().tolist())
sel_secs = st.sidebar.multiselect("Secretarias", sec_opts, default=sec_opts)

# Campo para filtrar por NÚMERO da conta digitando APENAS o prefixo (antes do hífen)
st.sidebar.subheader("Filtro por Conta (prefixo antes do hífen)")
prefix_input = st.sidebar.text_input("Digite apenas os números antes do hífen (ex.: 12345)", value="").strip()
prefix_digits = only_digits(prefix_input)

# Prepara coluna auxiliar com prefixo numérico
df["_conta_prefix"] = df["Conta"].apply(conta_prefix)

# Aplica filtros
df_f = df.copy()
if d_ini and d_fim:
    df_f = df_f[(df_f["Date"].dt.date >= d_ini) & (df_f["Date"].dt.date <= d_fim)]
if sel_secs:
    df_f = df_f[df_f["Secretaria"].isin(sel_secs)]
if prefix_digits:
    # filtra onde o prefixo numérico da conta começa com o que foi digitado
    df_f = df_f[df_f["_conta_prefix"].str.startswith(prefix_digits)]

if df_f.empty:
    st.warning("Nenhum dado após aplicar os filtros.")
    st.stop()

# =============== MÉTRICA SUPERIOR ===============
colA, colB = st.columns([1,1])
with colA:
    st.metric("Saldo Total (filtros)", brl(df_f["Saldo Bancario"].sum()))
with colB:
    if df_f["Date"].notna().any():
        st.caption(f"Período: {df_f['Date'].min().date().strftime('%d/%m/%Y')} → {df_f['Date'].max().date().strftime('%d/%m/%Y')}")

# =============== GRÁFICO POR SECRETARIA ===============
st.subheader("Saldo por Secretaria")

by_sec = (
    df_f.groupby("Secretaria", as_index=False)["Saldo Bancario"]
        .sum()
        .rename(columns={"Saldo Bancario": "Saldo"})
        .sort_values("Saldo", ascending=False)
        .reset_index(drop=True)
)

# Texto das barras em BRL; eixo y com prefixo R$
fig = px.bar(by_sec, x="Secretaria", y="Saldo", text=by_sec["Saldo"].map(brl))
fig.update_traces(textposition="outside")
fig.update_layout(
    yaxis_title="Saldo (R$)",
    xaxis_title="Secretaria",
    bargap=0.3,
    yaxis_tickprefix="R$ ",
    yaxis_tickformat=",.2f",
    uniformtext_minsize=8,
    uniformtext_mode="hide",
)
st.plotly_chart(fig, use_container_width=True)

# =============== RELATÓRIO "IGUAL À PLANILHA" ===============
st.subheader("Relatório (igual à planilha)")

df_rel = df_f.copy()
df_rel["Data"] = df_rel["Date"].dt.strftime("%d/%m/%Y")
df_rel["Saldo Bancario (R$)"] = df_rel["Saldo Bancario"].map(brl)

# Mantém ordem das colunas da planilha + Data no final
cols_order = ["Conta","Nome da Conta","Secretaria","Banco","Tipo de Recurso","Saldo Bancario (R$)","Data"]
st.dataframe(df_rel[cols_order].sort_values(["Data","Secretaria","Conta"]), use_container_width=True)

# Observação do filtro
if prefix_digits:
    st.caption(f"Filtro de conta aplicado: prefixo **{prefix_digits}** (parte numérica antes do hífen).")
