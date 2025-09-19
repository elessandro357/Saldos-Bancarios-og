import io
from datetime import datetime
import pandas as pd
import plotly.express as px
import streamlit as st

# ---------------- CONFIG ----------------
st.set_page_config(page_title="Saldos BB 2025 - Resumos", layout="wide")
pd.options.display.float_format = lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ---------------- HELPERS ----------------
def brl(x) -> str:
    try:
        return f"R$ {float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)

def load_all_sheets(xlsx_bytes_or_path) -> pd.DataFrame:
    """
    Lê todas as abas do Excel. A data é extraída do nome da aba (formato dd-mm-aaaa).
    Retorna DF com colunas:
      ['Conta','Nome da Conta','Secretaria','Banco','Tipo de Recurso','Saldo Bancario','Date']
    """
    xls = pd.ExcelFile(xlsx_bytes_or_path)
    frames = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        # Padroniza nomes
        df.columns = [c.strip() for c in df.columns]
        # Verifica colunas necessárias
        expected = {"Conta", "Nome da Conta", "Secretaria", "Banco", "Tipo de Recurso", "Saldo Bancario"}
        missing = expected.difference(df.columns)
        if missing:
            raise ValueError(f"Aba '{sheet}' sem colunas esperadas: {missing}")

        # Converte data a partir do nome da aba
        try:
            d = pd.to_datetime(sheet, format="%d-%m-%Y", dayfirst=True)
        except Exception:
            # Se não bater o formato, tenta outros (fallback)
            d = pd.to_datetime(sheet, dayfirst=True, errors="coerce")
        df["Date"] = d

        frames.append(df[["Conta","Nome da Conta","Secretaria","Banco","Tipo de Recurso","Saldo Bancario","Date"]])
    out = pd.concat(frames, ignore_index=True)
    # Tipagem
    out["Saldo Bancario"] = pd.to_numeric(out["Saldo Bancario"], errors="coerce")
    out["Secretaria"] = out["Secretaria"].astype(str).str.strip()
    out["Conta"] = out["Conta"].astype(str).str.strip()
    out["Nome da Conta"] = out["Nome da Conta"].astype(str).str.strip()
    out["Tipo de Recurso"] = out["Tipo de Recurso"].astype(str).str.strip()
    return out.dropna(subset=["Saldo Bancario"])

def to_excel_bytes(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            nm = name[:31]  # limite Excel
            df.to_excel(writer, index=False, sheet_name=nm)
            ws = writer.sheets[nm]
            for i, col in enumerate(df.columns):
                width = min(max(12, df[col].astype(str).map(len).max() if not df.empty else 12) + 2, 50)
                ws.set_column(i, i, width)
    return buf.getvalue()

# ---------------- INPUT ----------------
st.sidebar.header("Arquivo")
default_path = "/mnt/data/09 - Saldos BB 2025.xlsx"

up = st.sidebar.file_uploader("Envie o Excel (.xlsx) com as abas por data", type=["xlsx"])
use_default = st.sidebar.toggle("Usar arquivo padrão do ambiente", value=(up is None))

if up:
    source = up
elif use_default:
    try:
        with open(default_path, "rb") as f:
            source = io.BytesIO(f.read())
        st.sidebar.success("Arquivo padrão carregado.")
    except Exception:
        st.sidebar.error("Arquivo padrão não encontrado. Envie um .xlsx.")
        st.stop()
else:
    st.info("Envie a planilha .xlsx ou ative o arquivo padrão.")
    st.stop()

# ---------------- LOAD ----------------
try:
    df = load_all_sheets(source)
except Exception as e:
    st.error(f"Falha ao carregar planilha: {e}")
    st.stop()

if df.empty:
    st.warning("Sem dados.")
    st.stop()

# ---------------- FILTROS ----------------
st.title("Saldos por Secretaria, Conta e Tipo de Recurso (BB 2025)")

# Período
if df["Date"].notna().any():
    min_d = df["Date"].min().date()
    max_d = df["Date"].max().date()
    d_ini, d_fim = st.sidebar.date_input(
        "Período",
        value=(min_d, max_d),
        min_value=min_d,
        max_value=max_d
    )
else:
    d_ini, d_fim = None, None

# Secretaria / Conta / Tipo
sec_opts = sorted(df["Secretaria"].dropna().unique().tolist())
acc_opts = sorted(df["Conta"].dropna().unique().tolist())  # número
name_opts = sorted(df["Nome da Conta"].dropna().unique().tolist())  # nome
fund_opts = sorted(df["Tipo de Recurso"].dropna().unique().tolist())

sel_secs = st.sidebar.multiselect("Secretarias", sec_opts, default=sec_opts)
# Você pode filtrar por número de conta ou nome da conta (ou ambos)
sel_accs = st.sidebar.multiselect("Contas (nº)", acc_opts, default=acc_opts)
sel_names = st.sidebar.multiselect("Contas (nome)", name_opts, default=name_opts)
sel_funds = st.sidebar.multiselect("Tipo de Recurso", fund_opts, default=fund_opts)

# Aplica filtros
df_f = df.copy()
if d_ini and d_fim:
    mask = (df_f["Date"].dt.date >= d_ini) & (df_f["Date"].dt.date <= d_fim)
    df_f = df_f[mask]
if sel_secs:
    df_f = df_f[df_f["Secretaria"].isin(sel_secs)]
if sel_accs:
    df_f = df_f[df_f["Conta"].isin(sel_accs)]
if sel_names:
    df_f = df_f[df_f["Nome da Conta"].isin(sel_names)]
if sel_funds:
    df_f = df_f[df_f["Tipo de Recurso"].isin(sel_funds)]

if df_f.empty:
    st.warning("Nenhum dado após aplicar os filtros.")
    st.stop()

# ---------------- MÉTRICAS ----------------
colA, colB, colC = st.columns(3)
with colA:
    st.metric("Saldo Total (filtros)", brl(df_f["Saldo Bancario"].sum()))
with colB:
    st.metric("Secretarias", df_f["Secretaria"].nunique())
with colC:
    if df_f["Date"].notna().any():
        st.caption(f"{df_f['Date'].min().date().strftime('%d/%m/%Y')} → {df_f['Date'].max().date().strftime('%d/%m/%Y')}")

# ---------------- RESUMOS ----------------
st.subheader("1) Saldo por Secretaria")
by_sec = (df_f.groupby("Secretaria", as_index=False)["Saldo Bancario"].sum()
          .rename(columns={"Saldo Bancario": "Saldo"}).sort_values("Saldo", ascending=False))
fig1 = px.bar(by_sec, x="Secretaria", y="Saldo", text="Saldo")
fig1.update_traces(texttemplate="%{text:.2f}", textposition="outside")
fig1.update_layout(yaxis_title="Saldo (R$)", xaxis_title="Secretaria", bargap=0.3)
st.plotly_chart(fig1, use_container_width=True)
st.dataframe(by_sec.assign(Saldo=by_sec["Saldo"].map(brl)), use_container_width=True)

st.subheader("2) Saldo por Conta (nº)")
by_acc = (df_f.groupby("Conta", as_index=False)["Saldo Bancario"].sum()
          .rename(columns={"Saldo Bancario": "Saldo"}).sort_values("Saldo", ascending=False))
fig2 = px.bar(by_acc, x="Conta", y="Saldo", text="Saldo")
fig2.update_traces(texttemplate="%{text:.2f}", textposition="outside")
fig2.update_layout(yaxis_title="Saldo (R$)", xaxis_title="Conta", bargap=0.3)
st.plotly_chart(fig2, use_container_width=True)
st.dataframe(by_acc.assign(Saldo=by_acc["Saldo"].map(brl)), use_container_width=True)

st.subheader("3) Saldo por Nome da Conta")
by_acc_name = (df_f.groupby("Nome da Conta", as_index=False)["Saldo Bancario"].sum()
               .rename(columns={"Saldo Bancario": "Saldo"}).sort_values("Saldo", ascending=False))
fig2n = px.bar(by_acc_name, x="Nome da Conta", y="Saldo", text="Saldo")
fig2n.update_traces(texttemplate="%{text:.2f}", textposition="outside")
fig2n.update_layout(yaxis_title="Saldo (R$)", xaxis_title="Nome da Conta", bargap=0.3)
st.plotly_chart(fig2n, use_container_width=True)
st.dataframe(by_acc_name.assign(Saldo=by_acc_name["Saldo"].map(brl)), use_container_width=True)

st.subheader("4) Saldo por Tipo de Recurso")
by_fund = (df_f.groupby("Tipo de Recurso", as_index=False)["Saldo Bancario"].sum()
           .rename(columns={"Saldo Bancario": "Saldo"}).sort_values("Saldo", ascending=False))
fig3 = px.bar(by_fund, x="Tipo de Recurso", y="Saldo", text="Saldo")
fig3.update_traces(texttemplate="%{text:.2f}", textposition="outside")
fig3.update_layout(yaxis_title="Saldo (R$)", xaxis_title="Tipo de Recurso", bargap=0.3)
st.plotly_chart(fig3, use_container_width=True)
st.dataframe(by_fund.assign(Saldo=by_fund["Saldo"].map(brl)), use_container_width=True)

st.subheader("5) Cruzado: Secretaria × Conta × Tipo de Recurso")
cube = (df_f.groupby(["Secretaria","Conta","Nome da Conta","Tipo de Recurso"], as_index=False)["Saldo Bancario"].sum()
        .rename(columns={"Saldo Bancario": "Saldo"})
        .sort_values(["Secretaria","Conta","Tipo de Recurso"]))
st.dataframe(cube, use_container_width=True)

# ---------------- EVOLUÇÃO TEMPORAL ----------------
if df_f["Date"].notna().any():
    st.subheader("6) Evolução Temporal (Total por Dia)")
    by_day = (df_f.groupby("Date", as_index=False)["Saldo Bancario"].sum()
              .rename(columns={"Saldo Bancario": "Saldo"}).sort_values("Date"))
    fig4 = px.line(by_day, x="Date", y="Saldo", markers=True)
    fig4.update_layout(xaxis_title="Data", yaxis_title="Saldo (R$)")
    st.plotly_chart(fig4, use_container_width=True)

    st.subheader("7) Evolução por Secretaria (Empilhado)")
    by_day_sec = (df_f.groupby(["Date","Secretaria"], as_index=False)["Saldo Bancario"].sum()
                  .rename(columns={"Saldo Bancario": "Saldo"}))
    piv = by_day_sec.pivot(index="Date", columns="Secretaria", values="Saldo").fillna(0.0).sort_index()
    fig5 = px.area(piv, x=piv.index, y=piv.columns)
    fig5.update_layout(xaxis_title="Data", yaxis_title="Saldo (R$)")
    st.plotly_chart(fig5, use_container_width=True)

# ---------------- DETALHADO ----------------
st.subheader("8) Detalhado (com filtros)")
df_det = df_f.copy()
df_det["Data (dd/mm/aaaa)"] = df_det["Date"].dt.strftime("%d/%m/%Y")
df_det["Saldo"] = df_det["Saldo Bancario"].map(brl)
det_cols = ["Data (dd/mm/aaaa)","Secretaria","Conta","Nome da Conta","Banco","Tipo de Recurso","Saldo"]
st.dataframe(df_det[det_cols].sort_values(["Data (dd/mm/aaaa)","Secretaria","Conta"]), use_container_width=True)

# ---------------- EXPORT ----------------
st.subheader("Exportar")
exp_dept = by_sec.copy(); exp_dept["Saldo"] = exp_dept["Saldo"].round(2)
exp_acc  = by_acc.copy(); exp_acc["Saldo"]  = exp_acc["Saldo"].round(2)
exp_accn = by_acc_name.copy(); exp_accn["Saldo"] = exp_accn["Saldo"].round(2)
exp_fund = by_fund.copy(); exp_fund["Saldo"] = exp_fund["Saldo"].round(2)
exp_cube = cube.copy(); exp_cube["Saldo"] = exp_cube["Saldo"].round(2)
exp_day  = by_day if df_f["Date"].notna().any() else pd.DataFrame()

c1, c2 = st.columns(2)
with c1:
    all_csv = pd.concat([
        exp_dept.assign(__grupo__="Por Secretaria"),
        exp_acc.assign(__grupo__="Por Conta (nº)"),
        exp_accn.assign(__grupo__="Por Conta (nome)"),
        exp_fund.assign(__grupo__="Por Tipo de Recurso")
    ], ignore_index=True)
    st.download_button(
        "Baixar CSV (resumos)",
        data=all_csv.to_csv(index=False).encode("utf-8-sig"),
        file_name="saldos_resumos.csv",
        mime="text/csv"
    )
with c2:
    sheets = {
        "Por_Secretaria": exp_dept,
        "Por_Conta_Numero": exp_acc,
        "Por_Conta_Nome": exp_accn,
        "Por_Tipo_Recurso": exp_fund,
        "Cruzado": exp_cube,
        "Detalhado": df_det[det_cols]
    }
    if not exp_day.empty:
        sheets["Evolucao_Total_Dia"] = exp_day
    xls_bytes = to_excel_bytes(sheets)
    st.download_button(
        "Baixar Excel (abas)",
        data=xls_bytes,
        file_name="saldos_resumos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.caption("Este app assume o layout da sua planilha: abas por data e colunas padrão. Sem mapeamento manual.")
