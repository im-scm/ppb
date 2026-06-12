import re
import unicodedata
import pandas as pd
import streamlit as st
import plotly.express as px

# =========================================================
# CONFIG
# =========================================================
st.set_page_config(
    page_title="Cockpit Papel",
    layout="wide"
)

EXCEL_FILE = "Cockpit_Papel.xlsm"   # ajuste somente se o nome do arquivo no GitHub for diferente
SOURCE_SHEET = "Preços e Condições"  # base de dados
TITLE = "📊 Cockpit Papel - Dashboard Executivo"

st.title(TITLE)

# =========================================================
# HELPERS
# =========================================================
def normalize_text(value):
    """Normaliza texto para comparação robusta."""
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("utf-8")
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text

def parse_number(value):
    """Converte diferentes formatos numéricos para float."""
    if pd.isna(value):
        return None

    # já numérico
    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()

    if s == "" or s.lower() in {"nan", "none", "-", "--"}:
        return None

    # remove moeda e caracteres irrelevantes
    s = s.replace("R$", "").replace("$", "").replace("€", "")
    s = s.replace("%", "")
    s = s.replace(" ", "")

    # mantém só números, vírgula, ponto e sinal
    s = re.sub(r"[^0-9,.\-]", "", s)

    if s in {"", "-", ".", ","}:
        return None

    # casos com ponto e vírgula
    if "," in s and "." in s:
        # se a última vírgula vem depois do último ponto => formato BR 1.234,56
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            # formato US 1,234.56
            s = s.replace(",", "")
    elif "," in s:
        # assume vírgula decimal
        s = s.replace(".", "")
        s = s.replace(",", ".")
    else:
        # se tem muitos pontos, remove todos menos o último
        if s.count(".") > 1:
            parts = s.split(".")
            s = "".join(parts[:-1]) + "." + parts[-1]

    try:
        return float(s)
    except Exception:
        return None

def series_to_numeric(series):
    return series.apply(parse_number)

def detect_header_row(raw_df):
    """
    Encontra a linha que melhor se parece com o cabeçalho da base,
    procurando campos-chave.
    """
    targets = ["impress type", "supplier", "current price"]

    best_row = None
    best_score = -1

    max_rows = min(len(raw_df), 50)

    for i in range(max_rows):
        row_values = [normalize_text(v) for v in raw_df.iloc[i].tolist()]
        row_text = " | ".join(row_values)

        score = 0
        for t in targets:
            if t in row_text:
                score += 1

        if score > best_score:
            best_score = score
            best_row = i

    return best_row if best_score >= 2 else None

def find_column(df_columns, aliases):
    """
    Encontra a melhor coluna com base em aliases.
    """
    normalized_map = {col: normalize_text(col) for col in df_columns}

    for alias in aliases:
        alias_norm = normalize_text(alias)

        # match exato
        for col, col_norm in normalized_map.items():
            if col_norm == alias_norm:
                return col

        # match parcial
        for col, col_norm in normalized_map.items():
            if alias_norm in col_norm:
                return col

    return None

def build_canonical_dataframe(df):
    """
    Mapeia a planilha para o formato canônico do cockpit.
    """
    aliases = {
        "Impress Type": ["Impress Type", "Print Type"],
        "Width (mm)": ["Width (mm)", "Width mm", "Width", "Largura", "Largura (mm)"],
        "g/m2": ["g/m2", "g/m²", "Gramatura", "gsm"],
        "Supplier": ["Supplier", "Fornecedor"],
        "Currency": ["Currency", "Moeda"],
        "Current Price": ["Current Price", "Preço Atual", "Preco Atual", "CurrentPrice"],
        "Paper bonus (t)": ["Paper bonus (t)", "Paper bonus", "Bonus", "Bonus (t)"],
        "Lot (ton)": ["Lot (ton)", "Lot", "Lote", "Lote (ton)"],
        "TCO (R$/KG)": ["TCO (R$/KG)", "TCO R$/KG", "TCO KG", "TCO"],
        "TCO (R$/M2)": ["TCO (R$/M2)", "TCO R$/M2", "TCO M2"],
        "Payment Terms": ["Payment Terms", "Payment Term", "Prazo Pagamento"],
        "Working days": ["Working days", "Dias uteis", "Dias úteis"],
        "P.Value (R$/KG)": ["P.Value (R$/KG)", "P.Value R$/KG", "P Value (R$/KG)", "Present Value", "PV KG", "P.Value"],
        "P.Value (R$/M2)": ["P.Value (R$/M2)", "P.Value R$/M2", "P Value (R$/M2)", "PV M2"],
        "Última Atualização de Preço": ["Última Atualização de Preço", "Ultima Atualizacao de Preco", "Last Price Update", "Last Update"]
    }

    rename_map = {}
    for canonical_name, alias_list in aliases.items():
        original = find_column(df.columns, alias_list)
        if original:
            rename_map[original] = canonical_name

    df2 = df.rename(columns=rename_map).copy()

    # mantém apenas as colunas relevantes que existirem
    ordered_cols = [
        "Impress Type",
        "Width (mm)",
        "g/m2",
        "Supplier",
        "Currency",
        "Current Price",
        "Paper bonus (t)",
        "Lot (ton)",
        "TCO (R$/KG)",
        "TCO (R$/M2)",
        "Payment Terms",
        "Working days",
        "P.Value (R$/KG)",
        "P.Value (R$/M2)",
        "Última Atualização de Preço",
    ]

    existing_cols = [c for c in ordered_cols if c in df2.columns]
    df2 = df2[existing_cols].copy()

    # remove colunas históricas automaticamente (na prática já foram ignoradas),
    # mas isso protege caso alguma tenha passado
    keep_cols_set = set(ordered_cols)
    df2 = df2[[c for c in df2.columns if c in keep_cols_set]].copy()

    # conversão numérica robusta
    numeric_cols = [
        "Width (mm)",
        "g/m2",
        "Current Price",
        "Paper bonus (t)",
        "Lot (ton)",
        "TCO (R$/KG)",
        "TCO (R$/M2)",
        "Working days",
        "P.Value (R$/KG)",
        "P.Value (R$/M2)",
    ]

    for col in numeric_cols:
        if col in df2.columns:
            df2[col] = series_to_numeric(df2[col])

    # datas
    if "Última Atualização de Preço" in df2.columns:
        df2["Última Atualização de Preço"] = pd.to_datetime(
            df2["Última Atualização de Preço"],
            errors="coerce",
            dayfirst=True
        )

    # preenchimentos mínimos para manter o cockpit funcional
    if "TCO (R$/KG)" not in df2.columns and "Current Price" in df2.columns:
        df2["TCO (R$/KG)"] = df2["Current Price"]

    if "P.Value (R$/KG)" not in df2.columns and "TCO (R$/KG)" in df2.columns:
        df2["P.Value (R$/KG)"] = df2["TCO (R$/KG)"]

    if "TCO (R$/M2)" not in df2.columns and {"TCO (R$/KG)", "g/m2"}.issubset(df2.columns):
        df2["TCO (R$/M2)"] = df2["TCO (R$/KG)"] * (df2["g/m2"] / 1000.0)

    if "P.Value (R$/M2)" not in df2.columns and {"P.Value (R$/KG)", "g/m2"}.issubset(df2.columns):
        df2["P.Value (R$/M2)"] = df2["P.Value (R$/KG)"] * (df2["g/m2"] / 1000.0)

    # limpeza de linhas vazias essenciais
    essential = []
    if "Impress Type" in df2.columns:
        essential.append("Impress Type")
    if "Supplier" in df2.columns:
        essential.append("Supplier")

    if essential:
        df2 = df2.dropna(subset=essential)

    # se Current Price não existir, ao menos TCO precisa existir
    if "Current Price" in df2.columns:
        df2 = df2[df2["Current Price"].notna()]
    elif "TCO (R$/KG)" in df2.columns:
        df2 = df2[df2["TCO (R$/KG)"].notna()]

    # ordenação padrão
    sort_cols = [c for c in ["Impress Type", "Width (mm)", "g/m2", "Supplier", "Lot (ton)"] if c in df2.columns]
    if sort_cols:
        df2 = df2.sort_values(sort_cols, kind="stable")

    return df2

@st.cache_data
def load_data():
    # lê sem header para descobrir a linha do cabeçalho real
    raw = pd.read_excel(EXCEL_FILE, sheet_name=SOURCE_SHEET, header=None)

    header_row = detect_header_row(raw)
    if header_row is None:
        raise ValueError(
            "Não foi possível identificar automaticamente a linha de cabeçalho da aba 'Preços e Condições'."
        )

    df = pd.read_excel(EXCEL_FILE, sheet_name=SOURCE_SHEET, header=header_row)

    # limpeza básica dos cabeçalhos
    df.columns = [
        str(c).strip().replace("\n", " ")
        for c in df.columns
    ]
    df.columns = [re.sub(r"\s+", " ", c) for c in df.columns]

    # remove linhas totalmente vazias
    df = df.dropna(how="all")

    return build_canonical_dataframe(df)

# =========================================================
# LOAD
# =========================================================
try:
    df = load_data()
except Exception as e:
    st.error("❌ Falha ao ler e estruturar a base de dados.")
    st.code(str(e))
    st.stop()

if df.empty:
    st.warning("A base foi carregada, mas não há registros utilizáveis após a limpeza.")
    st.stop()

# =========================================================
# SIDEBAR = FILTROS APENAS
# =========================================================
st.sidebar.header("🎯 Filtros")

filtered = df.copy()

def multiselect_filter(df_in, column_name, label):
    if column_name not in df_in.columns:
        return df_in

    values = df_in[column_name].dropna().unique().tolist()

    # ordenação amigável
    try:
        values = sorted(values)
    except Exception:
        values = sorted(values, key=lambda x: str(x))

    selected = st.sidebar.multiselect(label, values, default=[])

    if selected:
        return df_in[df_in[column_name].isin(selected)]
    return df_in

# ordem conforme sua lógica de slicers
filtered = multiselect_filter(filtered, "Impress Type", "Impress Type")
filtered = multiselect_filter(filtered, "Width (mm)", "Width (mm)")
filtered = multiselect_filter(filtered, "g/m2", "g/m2")
filtered = multiselect_filter(filtered, "Supplier", "Supplier")
filtered = multiselect_filter(filtered, "Currency", "Currency")
filtered = multiselect_filter(filtered, "Lot (ton)", "Lot (ton)")

if filtered.empty:
    st.warning("Nenhum registro encontrado com os filtros selecionados.")
    st.stop()

# =========================================================
# KPIs
# =========================================================
def safe_min(series):
    s = pd.to_numeric(series, errors="coerce").dropna()
    return s.min() if not s.empty else None

def safe_max(series):
    s = pd.to_numeric(series, errors="coerce").dropna()
    return s.max() if not s.empty else None

def safe_best_supplier(df_in, metric_col):
    if metric_col not in df_in.columns:
        return None
    temp = df_in[[metric_col, "Supplier"]].copy()
    temp[metric_col] = pd.to_numeric(temp[metric_col], errors="coerce")
    temp = temp.dropna(subset=[metric_col, "Supplier"])
    if temp.empty:
        return None
    idx = temp[metric_col].idxmin()
    return temp.loc[idx, "Supplier"]

min_tco = safe_min(filtered["TCO (R$/KG)"]) if "TCO (R$/KG)" in filtered.columns else None
min_pv = safe_min(filtered["P.Value (R$/KG)"]) if "P.Value (R$/KG)" in filtered.columns else None
max_tco = safe_max(filtered["TCO (R$/KG)"]) if "TCO (R$/KG)" in filtered.columns else None
best_supplier = safe_best_supplier(filtered, "TCO (R$/KG)") if "TCO (R$/KG)" in filtered.columns else None

spread_pct = None
if min_tco is not None and max_tco is not None and min_tco != 0:
    spread_pct = ((max_tco / min_tco) - 1) * 100

k1, k2, k3, k4 = st.columns(4)
k1.metric("Melhor TCO (R$/KG)", f"{min_tco:,.4f}" if min_tco is not None else "N/A")
k2.metric("Melhor P.Value (R$/KG)", f"{min_pv:,.4f}" if min_pv is not None else "N/A")
k3.metric("Melhor Supplier", best_supplier if best_supplier else "N/A")
k4.metric("Spread TCO (%)", f"{spread_pct:,.1f}%" if spread_pct is not None else "N/A")

# =========================================================
# VISÕES PRINCIPAIS
# =========================================================
st.markdown("### Comparativo do mesmo Impress Type entre fornecedores")

left, right = st.columns(2)

with left:
    if "TCO (R$/KG)" in filtered.columns and "Supplier" in filtered.columns:
        chart_base = (
            filtered.groupby("Supplier", as_index=False)["TCO (R$/KG)"]
            .mean()
            .sort_values("TCO (R$/KG)", ascending=True)
        )

        fig_bar = px.bar(
            chart_base,
            x="Supplier",
            y="TCO (R$/KG)",
            text_auto=".4f",
            color="TCO (R$/KG)",
            color_continuous_scale="Blues",
            title="TCO médio por Supplier"
        )
        fig_bar.update_layout(showlegend=False)
        st.plotly_chart(fig_bar, use_container_width=True)

with right:
    scatter_x = "Lot (ton)" if "Lot (ton)" in filtered.columns else None
    scatter_y = "TCO (R$/KG)" if "TCO (R$/KG)" in filtered.columns else None
    if scatter_x and scatter_y:
        fig_scatter = px.scatter(
            filtered,
            x=scatter_x,
            y=scatter_y,
            color="Supplier" if "Supplier" in filtered.columns else None,
            hover_data=[c for c in ["Impress Type", "Width (mm)", "g/m2", "Currency", "Current Price"] if c in filtered.columns],
            title="Lot (ton) vs TCO (R$/KG)"
        )
        st.plotly_chart(fig_scatter, use_container_width=True)

# =========================================================
# TABELA PRINCIPAL (MODELO COCKPIT)
# =========================================================
st.markdown("### Tabela principal")

display_cols = [
    "Impress Type",
    "Width (mm)",
    "g/m2",
    "Supplier",
    "Currency",
    "Current Price",
    "Paper bonus (t)",
    "Lot (ton)",
    "TCO (R$/KG)",
    "TCO (R$/M2)",
    "Payment Terms",
    "Working days",
    "P.Value (R$/KG)",
    "P.Value (R$/M2)",
    "Última Atualização de Preço",
]

display_cols = [c for c in display_cols if c in filtered.columns]
table_df = filtered[display_cols].copy()

# formatação de data para exibição
if "Última Atualização de Preço" in table_df.columns:
    table_df["Última Atualização de Preço"] = table_df["Última Atualização de Preço"].dt.strftime("%d/%m/%Y")

st.dataframe(table_df, use_container_width=True, hide_index=True)

# =========================================================
# RESUMO POR IMPRESS TYPE
# =========================================================
if {"Impress Type", "Supplier", "TCO (R$/KG)"}.issubset(filtered.columns):
    st.markdown("### Resumo por Impress Type")

    summary = (
        filtered.groupby("Impress Type", as_index=False)
        .agg(
            Offers=("Supplier", "count"),
            Best_TCO_R_KG=("TCO (R$/KG)", "min"),
            Avg_TCO_R_KG=("TCO (R$/KG)", "mean"),
        )
        .sort_values(["Best_TCO_R_KG", "Impress Type"], ascending=[True, True])
    )

    st.dataframe(summary, use_container_width=True, hide_index=True)

# =========================================================
# RODAPÉ
# =========================================================
st.caption(
    "Observação: colunas históricas mensais foram ignoradas. "
    "Para comparação, o dashboard considera apenas 'Current Price' e os campos do modelo do cockpit."
)
