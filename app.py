import streamlit as st
import pandas as pd

# =============================
# CONFIG
# =============================
st.set_page_config(layout="wide")
st.title("📊 Cockpit Papel - Dashboard")

# =============================
# LOAD DATA
# =============================
@st.cache_data
def load_data():
    
    df = pd.read_excel("Cockpit_Papel.xlsm", sheet_name="Cockpit")

    # LIMPA NOMES DE COLUNAS
    df.columns = (
        df.columns
        .str.strip()              # remove espaços
        .str.replace("\n", " ")   # remove quebra de linha
        .str.replace("  ", " ")   # remove espaços duplos
    )
    return df

df = load_data()

# =============================
# CLEAN
# =============================
df = df.dropna(subset=["Supplier", "TCO (R$/KG)"])

# =============================
# SIDEBAR FILTERS
# =============================
st.sidebar.header("Filtros")

supplier = st.sidebar.multiselect(
    "Fornecedor",
    df["Supplier"].unique(),
    default=df["Supplier"].unique()
)

gramatura = st.sidebar.multiselect(
    "g/m²",
    sorted(df["g/m2"].unique()),
    default=sorted(df["g/m2"].unique())
)

width = st.sidebar.multiselect(
    "Width (mm)",
    sorted(df["Width\n (mm)"].unique()),
    default=sorted(df["Width\n (mm)"].unique())
)

df_filtered = df[
    (df["Supplier"].isin(supplier)) &
    (df["g/m2"].isin(gramatura)) &
    (df["Width\n (mm)"].isin(width))
]

# =============================
# KPIs
# =============================
col1, col2, col3, col4 = st.columns(4)

min_tco = df_filtered["TCO (R$/KG)"].min()
min_pv = df_filtered["P.Value (R$/KG)"].min()

best_supplier = df_filtered.loc[
    df_filtered["TCO (R$/KG)"].idxmin(), "Supplier"
]

col1.metric("💰 Menor TCO", f"{min_tco:.2f}")
col2.metric("📉 Melhor PV", f"{min_pv:.2f}")
col3.metric("🏆 Melhor fornecedor", best_supplier)
col4.metric("📊 Registros", len(df_filtered))

# =============================
# CHARTS
# =============================
st.subheader("📊 Comparação TCO por fornecedor")

chart = df_filtered.groupby("Supplier")["TCO (R$/KG)"].mean()
st.bar_chart(chart)

# =============================
# SCATTER (PV vs TCO)
# =============================
st.subheader("📉 PV vs TCO")

st.scatter_chart(
    df_filtered,
    x="TCO (R$/KG)",
    y="P.Value (R$/KG)",
)

# =============================
# TABLE
# =============================
st.subheader("📋 Dados")

st.dataframe(df_filtered, use_container_width=True)
