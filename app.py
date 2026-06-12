import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Cockpit Papel", layout="wide")
st.title("📊 Cockpit Papel - Dashboard Executivo")

# =========================================
# LOAD BASE (CORRETO)
# =========================================
@st.cache_data
def load_data():
    df = pd.read_excel("Cockpit_Papel.xlsm", sheet_name="Preços e Condições")

    # limpeza
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace("\n", " ")
        .str.replace(r"\s+", " ", regex=True)
    )

    return df

df = load_data()

# DEBUG (opcional)
st.sidebar.write("Colunas:", list(df.columns))

# =========================================
# NORMALIZA NOMES (ADAPTA AUTOMÁTICO)
# =========================================
def find_col(possible_names):
    for name in possible_names:
        for c in df.columns:
            if name.lower() in c.lower():
                return c
    return None

col_supplier = find_col(["supplier", "fornecedor"])
col_price = find_col(["price", "preço"])
col_width = find_col(["width", "largura"])
col_gram = find_col(["g/m2", "gramatura"])
col_currency = find_col(["currency", "moeda"])
col_lot = find_col(["lot", "lote"])

# =========================================
# VALIDAÇÃO
# =========================================
if col_supplier is None or col_price is None:
    st.error("❌ Não encontrei colunas essenciais (Supplier / Price)")
    st.stop()

# rename padrão
rename_map = {
    col_supplier: "Supplier",
    col_price: "Price"
}

if col_width: rename_map[col_width] = "Width"
if col_gram: rename_map[col_gram] = "Gramatura"
if col_currency: rename_map[col_currency] = "Currency"
if col_lot: rename_map[col_lot] = "Lot"

df = df.rename(columns=rename_map)

# =========================================
# ENGINEERING (AQUI É O SALTO DE QUALIDADE)
# =========================================

# Exemplo simples de TCO (ajuste se tiver fórmula mais completa)
df["TCO"] = df["Price"]

# Se quiser sofisticar depois:
# df["TCO"] = df["Price"] + freight + tax + etc

# =========================================
# FILTROS INTELIGENTES
# =========================================
st.sidebar.header("🎯 Filtros")

supplier = st.sidebar.multiselect(
    "Fornecedor",
    sorted(df["Supplier"].dropna().unique()),
    default=sorted(df["Supplier"].dropna().unique())
)

df_f = df[df["Supplier"].isin(supplier)]

if "Gramatura" in df_f.columns:
    gramatura = st.sidebar.multiselect(
        "Gramatura",
        sorted(df_f["Gramatura"].dropna().unique()),
        default=sorted(df_f["Gramatura"].dropna().unique())
    )
    df_f = df_f[df_f["Gramatura"].isin(gramatura)]

if "Width" in df_f.columns:
    width = st.sidebar.multiselect(
        "Width",
        sorted(df_f["Width"].dropna().unique()),
        default=sorted(df_f["Width"].dropna().unique())
    )
    df_f = df_f[df_f["Width"].isin(width)]

# =========================================
# KPIs
# =========================================
col1, col2, col3, col4 = st.columns(4)

min_tco = df_f["TCO"].min()
best_row = df_f.loc[df_f["TCO"].idxmin()]
best_supplier = best_row["Supplier"]

spread = ((df_f["TCO"].max() / min_tco) - 1) * 100

col1.metric("💰 Melhor TCO", f"{min_tco:,.2f}")
col2.metric("🏆 Melhor fornecedor", best_supplier)
col3.metric("📊 Spread", f"{spread:.1f}%")
col4.metric("📦 Registros", len(df_f))

# =========================================
# PIVOT (SUBSTITUI O EXCEL)
# =========================================
pivot = df_f.groupby("Supplier")["TCO"].mean().reset_index()

# =========================================
# GRÁFICOS
# =========================================
colA, colB = st.columns(2)

with colA:
    st.subheader("📊 TCO médio por fornecedor")

    fig = px.bar(
        pivot,
        x="Supplier",
        y="TCO",
        color="TCO",
        text_auto=True
    )

    st.plotly_chart(fig, use_container_width=True)

with colB:
    if "Lot" in df_f.columns:
        st.subheader("📈 TCO vs Lote")

        fig2 = px.scatter(
            df_f,
            x="Lot",
            y="TCO",
            color="Supplier"
        )

        st.plotly_chart(fig2, use_container_width=True)

# =========================================
# TABELA
# =========================================
st.subheader("📋 Base")

st.dataframe(df_f, use_container_width=True)

# =========================================
# INSIGHTS
# =========================================
st.subheader("🧠 Insights")

st.markdown(f"""
- Melhor fornecedor: **{best_supplier}**
- Melhor TCO: **{min_tco:,.2f}**
- Diferença competitiva: **{spread:.1f}%**
""")
