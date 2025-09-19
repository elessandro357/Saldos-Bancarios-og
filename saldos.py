import io
from datetime import datetime
import pandas as pd
import plotly.express as px
import streamlit as st

# ==================== CONFIG ====================
st.set_page_config(page_title="Saldos BB 2025 - Resumos Enxutos", layout="wide")
pd.options.display.float_format = lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ==================== HELPERS ====================
def brl(x) -> str:
    try:
        return f"R$ {float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)

def load_all_sheets(xlsx_bytes_or_path) -> pd.DataFrame:
    """
    Lê todas as abas (cada aba = um dia no formato dd-mm-aaaa) e retorna DF consolidado:
    ['Conta','Nome da Conta','Secretaria','Banco','Saldo Bancario','Date']
    """
    xls = pd.ExcelFile(xlsx_bytes_or_path)
    frames = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df.columns = [c.strip() for c in df.columns]

        expected = {"Conta", "Nome da Conta", "Secretaria", "Banco", "Saldo Bancario"}
        missing = expected.difference(df.columns)
        if missing:
            raise ValueError(f"Aba '{sheet}' sem colunas esperadas: {missing}")

        # Data a partir do nome da aba
        d = pd.to_datetime(sheet, format="%d-%m-%Y", dayfirst=True, errors="coerce")
        df["Date"] = d

        frames.append(df[["Conta","Nome da Conta","Secretaria","Banco","Saldo Bancario","Date"]])

    out = pd.concat(frames, ignore_index=True)
    out["Saldo Bancario"] = pd.to_numeric(out["Saldo Bancario"], errors="coerce")
    out["Secretaria"] = out["Secretaria"].astype(str).str.strip()
    out["Conta"] = out["Conta"].astype(str).str.strip()
    out["Nome da Conta"] = out["Nome da Conta"].astype(str).str.strip()
    return out.dropna(subset=["Saldo Bancario"])

def to_excel_bytes(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            nm = name[:31]
            df.to_excel(writer, index=False, sheet_name=nm)
            ws = writer.sheets[nm]
            for i, col in enumerate(df.columns):
                width = min(max(12, df[col].astype(str).map(len).max() if not df.empty else 12) + 2, 50)
                ws.set_column(i, i, width)
    return buf.getvalue()

# ==================== INPUT ====================
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

# ==================== LOAD ====================
try:
    df = load_all_sheets(source)
except Exception as e:
    st.error(f"Falha ao carregar planilha: {e}")
    st.stop()

if df.empty:
    st.warning("Sem dados.")
    st.stop()

# ==================== FILTROS ====================
st.title("Saldos: Secretaria, Conta (nº) e Nome da Conta")

# Período
if df["Date"].notna().any():
    min_d = df["Date"].min().date()
    max_d = df["Date"].max().date()
    d_ini, d_fim = st.sidebar.date_input("Período", value=(min_d, max_d), min_value=min_d, max_value=max_d)
else:
    d_ini, d_fim = None, None

sec_opts = sorted(df["Secretaria"].dropna().unique().tolist())
acc_opts = sorted(df["Conta"].dropna().unique().tolist())
name_opts = sorted(df["Nome da Conta"].dropna().unique().tolist())

sel_secs  = st.sidebar.multiselect("Secretarias", sec_opts, default=sec_opts)
sel_accs  = st.sidebar.multiselect("Contas (nº)", acc_opts, default=acc_opts)
sel_names = st.sidebar.multiselect("Contas (nome)", name_opts, default=name_opts)

df_f = df.copy()
if d_ini and d_fim:
    df_f = df_f[(df_f["Date"].dt.date >= d_ini) & (df_f["Date"].dt.date <= d_fim)]
if sel_secs:
    df_f = df_f[df_f["Secretaria"].isin(sel_secs)]
if sel_accs:
    df_f = df_f[df_f["Conta"].isin(sel_accs)]
if sel_names:
    df_f = df_f[df_f["Nome da Conta"].isin(sel_names)]

if df_f.empty:
    st.warning("Nenhum dado após aplicar os filtros.")
    st.stop()

# ==================== MÉTRICA SUPERIOR ====================
colA, colB = st.columns([1,1])
with colA:
    st.metric("Saldo Total (filtros)", brl(df_f["Saldo Bancario"].sum()))
with colB:
    if df_f["Date"].notna().any():
        st.caption(f"Período: {df_f['Date'].min().date().strftime('%d/%m/%Y')} → {df_f['Date'].max().date().strftime('%d/%m/%Y')}")

# ==================== 1) SALDO POR SECRETARIA ====================
st.subheader("Saldo por Secretaria")
by_sec = (df_f.groupby("Secretaria", as_index=False)["Saldo Bancario"].sum()
          .rename(columns={"Saldo Bancario": "Saldo"})
          .sort_values("Saldo", ascending=False))

fig_sec = px.bar(by_sec, x="Secretaria", y="Saldo", text="Saldo")
fig_sec.update_traces(texttemplate="%{text:.2f}", textposition="outside")
fig_sec.update_layout(yaxis_title="Saldo (R$)", xaxis_title="Secretaria", bargap=0.3)
st.plotly_chart(fig_sec, use_container_width=True)

st.dataframe(by_sec.assign(Saldo=by_sec["Saldo"].map(brl)), use_container_width=True)

# ==================== 2) SALDO POR CONTA (NÚMERO) ====================
st.subheader("Saldo por Conta (nº)")
by_acc = (df_f.groupby("Conta", as_index=False)["Saldo Bancario"].sum()
          .rename(columns={"Saldo Bancario": "Saldo"})
          .sort_values("Saldo", ascending=False))

fig_acc = px.bar(by_acc, x="Conta", y="Saldo", text="Saldo")
fig_acc.update_traces(texttemplate="%{text:.2f}", textposition="outside")
fig_acc.update_layout(yaxis_title="Saldo (R$)", xaxis_title="Conta (nº)", bargap=0.3)
st.plotly_chart(fig_acc, use_container_width=True)

st.dataframe(by_acc.assign(Saldo=by_acc["Saldo"].map(brl)), use_container_width=True)

# ==================== 3) SALDO POR NOME DA CONTA (TABELA) ====================
st.subheader("Saldo por Nome da Conta")
by_acc_name = (df_f.groupby("Nome da Conta", as_index=False)["Saldo Bancario"].sum()
               .rename(columns={"Saldo Bancario": "Saldo"})
               .sort_values("Saldo", ascending=False))

st.dataframe(by_acc_name.assign(Saldo=by_acc_name["Saldo"].map(brl)), use_container_width=True)

# ==================== EXPORTS ====================
st.subheader("Exportar (resumos desta tela)")
exp_dept = by_sec.copy();      exp_dept["Saldo"] = exp_dept["Saldo"].round(2)
exp_acc  = by_acc.copy();      exp_acc["Saldo"]  = exp_acc["Saldo"].round(2)
exp_name = by_acc_name.copy(); exp_name["Saldo"] = exp_name["Saldo"].round(2)

c1, c2 = st.columns(2)
with c1:
    csv_bytes = pd.concat([
        exp_dept.assign(__grupo__="Por Secretaria"),
        exp_acc.assign(__grupo__="Por Conta (nº)"),
        exp_name.assign(__grupo__="Por Conta (nome)")
    ], ignore_index=True).to_csv(index=False).encode("utf-8-sig")
    st.download_button("Baixar CSV (resumos)", data=csv_bytes, file_name="saldos_resumos.csv", mime="text/csv")

with c2:
    xlsx_bytes = to_excel_bytes({
        "Por_Secretaria": exp_dept,
        "Por_Conta_Num": exp_acc,
        "Por_Conta_Nome": exp_name
    })
    st.download_button("Baixar Excel (abas)", data=xlsx_bytes,
                       file_name="saldos_resumos.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Tela simplificada: apenas Secretaria, Conta (nº) e Nome da Conta.")
