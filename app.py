import streamlit as st
import pandas as pd
import plotly.express as px

# CONFIG
st.set_page_config(page_title="Cockpit Papel", layout="wide")
st.title("📊 Cockpit Papel - Dashboard Executivo")

# LOAD DATA
@st.cache_data
def load_data():
    df = pd.read_excel("Cockpit_Papel.xlsm", sheet_name="Cockpit")

    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace("\n", " ")
        .str.replace(r"\s+", " ", regex=True)
    )

    return df

df = load_data()

# =============================
# DETECÇÃO AUTOMÁTICA
# =============================
def find_col(keyword):
    cols = [c for c in df.columns if keyword.lower() in c.lower()]
    return cols[0] if cols else None

col_supplier = find_col("Supplier")
col_width = find_col("Width")
col_gramatura = find_col("g/m2")
col_tco = find_col("TCO")
col_pv = find_col("P.Value")

df = df.rename(columns={
    col_supplier: "Supplier",
    col_width: "Width",
    col_gramatura: "Gramatura",
    col_tco: "TCO",
    col_pv: "PV"
})

df = df.dropna(subset=["Supplier", "TCO"])

# =============================
# FILTROS
# =============================
st.sidebar.header("🎯 Filtros")

supplier_filter = st.sidebar.multiselect(
    "Fornecedor",
    sorted(df["Supplier"].unique()),
    default=sorted(df["Supplier"].unique())
)

df_f = df[df["Supplier"].isin(supplier_filter)]

gramatura_filter = st.sidebar.multiselect(
    "Gramatura",
    sorted(df_f["Gramatura"].dropna().unique()),
    default=sorted(df_f["Gramatura"].dropna().unique())
)

df_f = df_f[df_f["Gramatura"].isin(gramatura_filter)]

width_filter = st.sidebar.multiselect(
    "Width",
    sorted(df_f["Width"].dropna().unique()),
    default=sorted(df_f["Width"].dropna().unique())
)

df_f = df_f[df_f["Width"].isin(width_filter)]

# =============================
# KPIs
# =============================
col1, col2, col3, col4 = st.columns(4)

min_tco = df_f["TCO"].min()
min_pv = df_f["PV"].min()

best_row = df_f.loc[df_f["TCO"].idxmin()]
best_supplier = best_row["Supplier"]

spread = ((df_f["TCO"].max() / min_tco) - 1) * 100

col1.metric("💰 Melhor TCO", f"{min_tco:,.2f}")
col2.metric("📉 Melhor PV", f"{min_pv:,.2f}")
col3.metric("🏆 Fornecedor", best_supplier)
col4.metric("📊 Spread", f"{spread:.1f}%")

# =============================
# GRÁFICOS
# =============================
colA, colB = st.columns(2)

with colA:
    st.subheader("TCO por fornecedor")
    df_chart = df_f.groupby("Supplier")["TCO"].mean().reset_index()

    fig = px.bar(df_chart, x="Supplier", y="TCO", color="TCO")
    st.plotly_chart(fig, use_container_width=True)

with colB:
    st.subheader("PV vs TCO")

    fig2 = px.scatter(
        df_f,
        x="TCO",
        y="PV",
        color="Supplier",
        size="Gramatura",
        hover_data=["Width"],
    )

    st.plotly_chart(fig2, use_container_width=True)

# =============================
# CURVA DE LOTE
# =============================
if "Lot" in df_f.columns:
    st.subheader("Curva TCO vs Lote")

    fig3 = px.scatter(
        df_f,
        x="Lot",
        y="TCO",
        color="Supplier",
        trendline="ols"
    )

    st.plotly_chart(fig3, use_container_width=True)

# =============================
# TABELA
# =============================
st.subheader("Base filtrada")
st.dataframe(df_f, use_container_width=True)

# =============================
# INSIGHTS
# =============================
st.subheader("Insights")

st.markdown(f"""
- Melhor fornecedor: **{best_supplier}**
- Melhor TCO: **{min_tco:,.2f}**
- Spread: **{spread:.1f}%**
""")
