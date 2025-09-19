import io
from datetime import date
from typing import Optional, List

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ------------------ CONFIG ------------------
st.set_page_config(page_title="Saldos por Secretaria, Conta e Tipo de Recurso", layout="wide")
pd.options.display.float_format = lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ------------------ HELPERS ------------------
def brl(x) -> str:
    try:
        return f"R$ {float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)

def parse_dates(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=True, infer_datetime_format=True)

def sign_by_type(df: pd.DataFrame, amount_col: str, type_col: Optional[str],
                 credit_tokens: List[str], debit_tokens: List[str]) -> pd.Series:
    vals = pd.to_numeric(df[amount_col], errors="coerce")
    if type_col is None:
        return vals  # já vem com sinal, ou tudo positivo
    t = df[type_col].astype(str).str.strip().str.upper()
    credit_set = {s.strip().upper() for s in credit_tokens if s.strip()}
    debit_set  = {s.strip().upper() for s in debit_tokens  if s.strip()}
    signed = vals.copy()
    is_credit = t.isin(credit_set)
    is_debit  = t.isin(debit_set)
    signed[is_credit] = vals[is_credit].abs()
    signed[is_debit]  = -vals[is_debit].abs()
    return signed

def to_excel_bytes(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=name[:31])
            ws = writer.sheets[name[:31]]
            for i, col in enumerate(df.columns):
                width = min(max(12, df[col].astype(str).map(len).max() if not df.empty else 12) + 2, 50)
                ws.set_column(i, i, width)
    return buf.getvalue()

# ------------------ SIDEBAR: INPUT ------------------
st.sidebar.header("Dados")
default_path = "/mnt/data/Saldos BB 2025 - 11-09.xlsx"
uploaded = st.sidebar.file_uploader("Planilha Excel (.xlsx)", type=["xlsx"])
use_default = False
if not uploaded:
    use_default = st.sidebar.toggle("Usar arquivo padrão do ambiente", value=False)
    if use_default:
        try:
            with open(default_path, "rb") as f:
                uploaded = io.BytesIO(f.read())
            st.sidebar.success("Arquivo padrão carregado.")
        except Exception:
            st.sidebar.warning("Arquivo padrão não encontrado. Envie um .xlsx.")

if not uploaded:
    st.info("Envie a planilha .xlsx (ou habilite o padrão).")
    st.stop()

# Descobrir abas
try:
    xls = pd.ExcelFile(uploaded)
    sheets = xls.sheet_names
except Exception as e:
    st.error(f"Falha ao abrir o Excel: {e}")
    st.stop()

sheet = st.sidebar.selectbox("Aba da planilha", sheets, index=0)
try:
    df_raw = pd.read_excel(uploaded, sheet_name=sheet)
except Exception as e:
    st.error(f"Falha ao ler a aba '{sheet}': {e}")
    st.stop()

if df_raw.empty:
    st.warning("A aba selecionada está vazia.")
    st.stop()

cols = list(df_raw.columns)

st.sidebar.subheader("Mapeamento de Colunas")
col_secretaria = st.sidebar.selectbox("Secretaria*", options=cols)
# Conta pode ser número OU nome; você escolhe qual usar
col_conta = st.sidebar.selectbox("Conta (nº ou nome)*", options=["<Nenhuma>"] + cols, index=0)
col_tipo_recurso = st.sidebar.selectbox("Tipo de Recurso (livre/vinculado)*", options=["<Nenhuma>"] + cols, index=0)
col_valor = st.sidebar.selectbox("Valor*", options=cols)
col_data = st.sidebar.selectbox("Data (opcional)", options=["<Nenhuma>"] + cols, index=0)
col_tipo_lcto = st.sidebar.selectbox("Tipo Lançamento (C/D) (opcional)", options=["<Nenhuma>"] + cols, index=0)

has_date = col_data != "<Nenhuma>"
has_type = col_tipo_lcto != "<Nenhuma>"
has_conta = col_conta != "<Nenhuma>"
has_tipo_recurso = col_tipo_recurso != "<Nenhuma>"

credit_tokens = []
debit_tokens = []
if has_type:
    with st.sidebar.expander("Mapeamento C/D", expanded=False):
        st.caption("Informe como sua planilha marca os lançamentos.")
        credit_tokens = [s for s in st.text_input("Crédito (ex.: C, CR, CREDITO)", value="C, CR, CREDITO").split(",")]
        debit_tokens  = [s for s in st.text_input("Débito  (ex.: D, DB, DEBITO)", value="D, DB, DEBITO").split(",")]

# ------------------ PREPARO ------------------
df = df_raw.copy()

# Datas
if has_date:
    df["_date"] = parse_dates(df[col_data])
else:
    df["_date"] = pd.NaT

# Valor assinado
df["_amount"] = sign_by_type(df, amount_col=col_valor,
                             type_col=(col_tipo_lcto if has_type else None),
                             credit_tokens=credit_tokens, debit_tokens=debit_tokens)

# Chaves
df["_dept"] = df[col_secretaria].astype(str).str.strip()
if has_conta:
    df["_account"] = df[col_conta].astype(str).str.strip()
else:
    df["_account"] = "Sem conta"

if has_tipo_recurso:
    df["_fundtype"] = df[col_tipo_recurso].astype(str).str.strip().str.lower().map({
        "livre": "Livre", "vinculado": "Vinculado", "vinculada": "Vinculado", "nao livre": "Vinculado",
        "não livre": "Vinculado", "nao-vinculado": "Livre", "não-vinculado": "Livre"
    }).fillna(df[col_tipo_recurso].astype(str).str.strip())
else:
    df["_fundtype"] = "Não informado"

# Limpeza
df = df[~df["_dept"].isna() & df["_dept"].ne("")]
df = df[~df["_amount"].isna()]

# ------------------ FILTROS ------------------
st.sidebar.header("Filtros")
dept_opts = sorted(df["_dept"].unique().tolist())
sel_depts = st.sidebar.multiselect("Secretarias", dept_opts, default=dept_opts)

acct_opts = sorted(df["_account"].unique().tolist())
sel_accts = st.sidebar.multiselect("Contas", acct_opts, default=acct_opts)

fund_opts = sorted(df["_fundtype"].unique().tolist())
sel_funds = st.sidebar.multiselect("Tipo de Recurso", fund_opts, default=fund_opts)

if has_date and df["_date"].notna().any():
    min_d = pd.to_datetime(df["_date"].min()).date()
    max_d = pd.to_datetime(df["_date"].max()).date()
    d_ini, d_fim = st.sidebar.date_input("Período", value=(min_d, max_d), min_value=min_d, max_value=max_d)
else:
    d_ini, d_fim = None, None

df_f = df.copy()
if sel_depts:
    df_f = df_f[df_f["_dept"].isin(sel_depts)]
if sel_accts:
    df_f = df_f[df_f["_account"].isin(sel_accts)]
if sel_funds:
    df_f = df_f[df_f["_fundtype"].isin(sel_funds)]
if has_date and d_ini and d_fim:
    mask = (df_f["_date"].dt.date >= d_ini) & (df_f["_date"].dt.date <= d_fim)
    df_f = df_f[mask]

if df_f.empty:
    st.warning("Nenhum dado após aplicar os filtros.")
    st.stop()

# ------------------ TÍTULO ------------------
st.title("Saldos por Secretaria, Conta e Tipo de Recurso")

# ------------------ RESUMOS ------------------
col1, col2, col3 = st.columns(3)
with col1:
    total_geral = df_f["_amount"].sum()
    st.metric("Saldo Total (filtros)", brl(total_geral))
with col2:
    st.metric("Secretarias", df_f["_dept"].nunique())
with col3:
    if has_date and df_f["_date"].notna().any():
        st.caption(f"Período: **{df_f['_date'].min().date().strftime('%d/%m/%Y')}** a **{df_f['_date'].max().date().strftime('%d/%m/%Y')}**")
    else:
        st.caption("Período: não informado")

# --- 1) SALDO POR SECRETARIA ---
st.subheader("Saldo por Secretaria")
by_dept = df_f.groupby("_dept", as_index=False)["_amount"].sum().rename(columns={"_dept": "Secretaria", "_amount": "Saldo"})
by_dept = by_dept.sort_values("Saldo", ascending=False).reset_index(drop=True)

fig_dept = px.bar(by_dept, x="Secretaria", y="Saldo", text="Saldo", title=None)
fig_dept.update_traces(texttemplate="%{text:.2f}", textposition="outside")
fig_dept.update_layout(yaxis_title="Saldo (R$)", xaxis_title="Secretaria", bargap=0.3)
st.plotly_chart(fig_dept, use_container_width=True)
st.dataframe(by_dept.assign(Saldo=by_dept["Saldo"].map(brl)), use_container_width=True)

# --- 2) SALDO POR CONTA ---
st.subheader("Saldo por Conta")
by_acct = df_f.groupby("_account", as_index=False)["_amount"].sum().rename(columns={"_account": "Conta", "_amount": "Saldo"})
by_acct = by_acct.sort_values("Saldo", ascending=False).reset_index(drop=True)

fig_acct = px.bar(by_acct, x="Conta", y="Saldo", text="Saldo", title=None)
fig_acct.update_traces(texttemplate="%{text:.2f}", textposition="outside")
fig_acct.update_layout(yaxis_title="Saldo (R$)", xaxis_title="Conta", bargap=0.3)
st.plotly_chart(fig_acct, use_container_width=True)
st.dataframe(by_acct.assign(Saldo=by_acct["Saldo"].map(brl)), use_container_width=True)

# --- 3) SALDO POR TIPO DE RECURSO ---
st.subheader("Saldo por Tipo de Recurso")
by_fund = df_f.groupby("_fundtype", as_index=False)["_amount"].sum().rename(columns={"_fundtype": "Tipo de Recurso", "_amount": "Saldo"})
by_fund = by_fund.sort_values("Saldo", ascending=False).reset_index(drop=True)

fig_fund = px.bar(by_fund, x="Tipo de Recurso", y="Saldo", text="Saldo", title=None)
fig_fund.update_traces(texttemplate="%{text:.2f}", textposition="outside")
fig_fund.update_layout(yaxis_title="Saldo (R$)", xaxis_title="Tipo de Recurso", bargap=0.3)
st.plotly_chart(fig_fund, use_container_width=True)
st.dataframe(by_fund.assign(Saldo=by_fund["Saldo"].map(brl)), use_container_width=True)

# --- 4) CUBO CRUZADO: Secretaria x Conta x Tipo de Recurso ---
st.subheader("Cruzado: Secretaria x Conta x Tipo de Recurso")
cube = df_f.groupby(["_dept", "_account", "_fundtype"], as_index=False)["_amount"].sum()
cube = cube.rename(columns={"_dept": "Secretaria", "_account": "Conta", "_fundtype": "Tipo de Recurso", "_amount": "Saldo"})
st.dataframe(cube.sort_values(["Secretaria", "Conta", "Tipo de Recurso"]), use_container_width=True)

# Evolução mensal (se tiver data)
if has_date and df_f["_date"].notna().any():
    st.subheader("Evolução Mensal do Saldo (somatório de lançamentos)")
    df_f["_ym"] = df_f["_date"].dt.to_period("M").dt.to_timestamp()
    monthly = df_f.groupby(["_ym", "_dept"], as_index=False)["_amount"].sum()
    piv = monthly.pivot(index="_ym", columns="_dept", values="_amount").fillna(0.0).sort_index()
    fig_m = px.area(piv, x=piv.index, y=piv.columns)
    fig_m.update_layout(xaxis_title="Mês", yaxis_title="Saldo (R$)")
    st.plotly_chart(fig_m, use_container_width=True)

# ------------------ EXPORTS ------------------
st.subheader("Exportar Relatórios")
export_dept = by_dept.copy(); export_dept["Saldo"] = export_dept["Saldo"].round(2)
export_acct = by_acct.copy(); export_acct["Saldo"] = export_acct["Saldo"].round(2)
export_fund = by_fund.copy(); export_fund["Saldo"] = export_fund["Saldo"].round(2)
export_cube = cube.copy();    export_cube["Saldo"] = export_cube["Saldo"].round(2)

c1, c2 = st.columns(2)
with c1:
    st.download_button(
        "Baixar CSV (todos os resumos)",
        data=pd.concat([
            export_dept.assign(__grupo__="Por Secretaria"),
            export_acct.assign(__grupo__="Por Conta"),
            export_fund.assign(__grupo__="Por Tipo de Recurso"),
        ], ignore_index=True).to_csv(index=False).encode("utf-8-sig"),
        file_name="saldos_resumos.csv",
        mime="text/csv"
    )
with c2:
    xls_bytes = to_excel_bytes({
        "Por_Secretaria": export_dept,
        "Por_Conta": export_acct,
        "Por_Tipo_Recurso": export_fund,
        "Cruzado": export_cube
    })
    st.download_button(
        "Baixar Excel (abas)",
        data=xls_bytes,
        file_name="saldos_resumos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.caption("Mapeie Secretaria, Conta e Tipo de Recurso na barra lateral. Se houver coluna de C/D, informe os tokens para assinar valores corretamente.")
