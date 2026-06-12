import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Cockpit Papel", layout="wide")
st.title("📊 Cockpit Papel - Dashboard Executivo")

# =========================================
# LOAD DATA
# =========================================
@st.cache_data
def load_data():
    df = pd.read_excel("Cockpit_Papel.xlsm", sheet_name="Cockpit")

    # limpeza robusta
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace("\n", " ")
        .str.replace(r"\s+", " ", regex=True)
    )

    return df

df = load_data()

# =========================================
# DEBUG (IMPORTANTE)
# =========================================
st.sidebar.write("🔍 Colunas detectadas:")
st.sidebar.write(list(df.columns))

# =========================================
# DETECÇÃO SEGURA
# =========================================
def safe_find(keyword):
    for c in df.columns:
        if keyword.lower() in c.lower():
            return c
    return None

col_supplier = safe_find("Supplier")
col_width = safe_find("Width")
col_gramatura = safe_find("g/m2")
col_tco = safe_find("TCO")
col_pv = safe_find("P.Value")

# =========================================
# VALIDAÇÃO (IMPEDIR QUEBRA)
# =========================================
if col_supplier is None or col_tco is None:
    st.error("❌ Não foi possível identificar as colunas principais.")
    st.write("👉 Verifique os nomes exibidos na sidebar.")
    st.stop()

# =========================================
# PADRONIZAÇÃO
# =========================================
rename_map = {}

rename_map[col_supplier] = "Supplier"
rename_map[col_tco] = "TCO"

if col_width: rename_map[col_width] = "Width"
if col_gramatura: rename_map[col_gramatura] = "Gramatura"
if col_pv: rename_map[col_pv] = "PV"

df = df.rename(columns=rename_map)

# =========================================
# FILTRO BASE
# =========================================
df = df[df["Supplier"].notna()]
df = df[df["TCO"].notna()]

# =========================================
# SIDEBAR FILTROS
# =========================================
st.sidebar.header("🎯 Filtros")

supplier_filter = st.sidebar.multiselect(
    "Fornecedor",
    sorted(df["Supplier"].unique()),
    default=sorted(df["Supplier"].unique())
)

df_f = df[df["Supplier"].isin(supplier_filter)]

if "Gramatura" in df_f.columns:
    gramatura_filter = st.sidebar.multiselect(
        "Gramatura",
        sorted(df_f["Gramatura"].dropna().unique()),
        default=sorted(df_f["Gramatura"].dropna().unique())
    )
    df_f = df_f[df_f["Gramatura"].isin(gramatura_filter)]

if "Width" in df_f.columns:
    width_filter = st.sidebar.multiselect(
        "Width",
        sorted(df_f["Width"].dropna().unique()),
        default=sorted(df_f["Width"].dropna().unique())
    )
    df_f = df_f[df_f["Width"].isin(width_filter)]

# =========================================
# KPIs
# =========================================
col1, col2, col3, col4 = st.columns(4)

min_tco = df_f["TCO"].min()

if "PV" in df_f.columns:
    min_pv = df_f["PV"].min()
else:
    min_pv = None

best_row = df_f.loc[df_f["TCO"].idxmin()]
best_supplier = best_row["Supplier"]

spread = ((df_f["TCO"].max() / min_tco) - 1) * 100

col1.metric("💰 Melhor TCO", f"{min_tco:,.2f}")

if min_pv:
    col2.metric("📉 Melhor PV", f"{min_pv:,.2f}")
else:
    col2.metric("📉 PV", "N/A")

col3.metric("🏆 Fornecedor", best_supplier)
col4.metric("📊 Spread", f"{spread:.1f}%")

# =========================================
# GRÁFICOS
# =========================================
colA, colB = st.columns(2)

with colA:
    st.subheader("TCO por fornecedor")

    df_chart = df_f.groupby("Supplier")["TCO"].mean().reset_index()

    fig = px.bar(df_chart, x="Supplier", y="TCO", color="TCO")
    st.plotly_chart(fig, use_container_width=True)

with colB:
    st.subheader("PV vs TCO")

    if "PV" in df_f.columns:
        fig2 = px.scatter(
            df_f,
            x="TCO",
            y="PV",
            color="Supplier",
            hover_data=df_f.columns
        )
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("PV não disponível")

# =========================================
# CURVA DE LOTE
# =========================================
if "Lot" in df_f.columns:
    st.subheader("Curva TCO vs Lote")

    fig3 = px.scatter(
        df_f,
        x="Lot",
        y="TCO",
        color="Supplier"
    )

    st.plotly_chart(fig3, use_container_width=True)

# =========================================
# TABELA
# =========================================
st.subheader("Base filtrada")
st.dataframe(df_f, use_container_width=True)

# =========================================
# INSIGHTS
# =========================================
st.subheader("Insights")

st.markdown(f"""
- Melhor fornecedor: **{best_supplier}**
- Melhor TCO: **{min_tco:,.2f}**
- Spread: **{spread:.1f}%**
""")
