import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Cockpit Papel", layout="wide")
st.title("📊 Cockpit Papel - Dashboard Executivo")

# =========================================
# FUNÇÃO INTELIGENTE DE LEITURA
# =========================================
@st.cache_data
def load_data():

    raw = pd.read_excel("Cockpit_Papel.xlsm", sheet_name="Preços e Condições", header=None)

    # 🔍 encontra a linha que contém "Supplier" ou "Fornecedor"
    header_row = None

    for i, row in raw.iterrows():
        row_str = " ".join([str(x) for x in row.values])

        if "supplier" in row_str.lower() or "fornecedor" in row_str.lower():
            header_row = i
            break

    if header_row is None:
        st.error("❌ Não encontrei o header da tabela")
        st.stop()

    # ✅ recria dataframe corretamente
    df = pd.read_excel(
        "Cockpit_Papel.xlsm",
        sheet_name="Preços e Condições",
        header=header_row
    )

    # limpeza
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace("\n", " ")
        .str.replace(r"\s+", " ", regex=True)
    )

    return df

df = load_data()

# DEBUG
st.sidebar.write("Colunas:", list(df.columns))

# =========================================
# DETECÇÃO FLEXÍVEL
# =========================================
def find_col(possible_names):
    for name in possible_names:
        for c in df.columns:
            if name.lower() in c.lower():
                return c
    return None

col_supplier = find_col(["supplier", "fornecedor"])
col_price = find_col(["price", "preço", "valor"])
col_width = find_col(["width", "largura"])
col_gram = find_col(["g/m2", "gramatura"])
col_lot = find_col(["lot", "lote"])
col_currency = find_col(["currency", "moeda"])

# ✅ validação real
if col_supplier is None or col_price is None:
    st.error("❌ Não encontrei Supplier ou Price. Veja sidebar.")
    st.stop()

# padroniza
df = df.rename(columns={
    col_supplier: "Supplier",
    col_price: "Price",
    **({col_width: "Width"} if col_width else {}),
    **({col_gram: "Gramatura"} if col_gram else {}),
    **({col_lot: "Lot"} if col_lot else {}),
    **({col_currency: "Currency"} if col_currency else {})
})

# =========================================
# LIMPEZA
# =========================================
df = df[df["Supplier"].notna()]
df = df[df["Price"].notna()]

# =========================================
# ENGINEERING (TCO BASE)
# =========================================
df["TCO"] = df["Price"]

# =========================================
# FILTROS
# =========================================
st.sidebar.header("🎯 Filtros")

supplier = st.sidebar.multiselect(
    "Fornecedor",
    sorted(df["Supplier"].unique()),
    default=sorted(df["Supplier"].unique())
)

df_f = df[df["Supplier"].isin(supplier)]

if "Gramatura" in df_f.columns:
    gram = st.sidebar.multiselect(
        "Gramatura",
        sorted(df_f["Gramatura"].dropna().unique()),
        default=sorted(df_f["Gramatura"].dropna().unique())
    )
    df_f = df_f[df_f["Gramatura"].isin(gram)]

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
col2.metric("🏆 Fornecedor", best_supplier)
col3.metric("📊 Spread", f"{spread:.1f}%")
col4.metric("📦 Registros", len(df_f))

# =========================================
# PIVOT (RECRIADO)
# =========================================
pivot = df_f.groupby("Supplier")["TCO"].mean().reset_index()

colA, colB = st.columns(2)

with colA:
    st.subheader("TCO médio por fornecedor")
    fig = px.bar(pivot, x="Supplier", y="TCO", color="TCO")
    st.plotly_chart(fig, use_container_width=True)

with colB:
    if "Lot" in df_f.columns:
        st.subheader("TCO vs Lote")
        fig2 = px.scatter(df_f, x="Lot", y="TCO", color="Supplier")
        st.plotly_chart(fig2, use_container_width=True)

# =========================================
# TABELA
# =========================================
st.subheader("Base")
st.dataframe(df_f, use_container_width=True)
