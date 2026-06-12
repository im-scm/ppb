import re
import unicodedata
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

# =========================================================
# CONFIG
# =========================================================
st.set_page_config(
    page_title="Cockpit Papel",
    layout="wide"
)

EXCEL_FILE = "Cockpit_Papel.xlsm"          # ajuste apenas se o nome do arquivo estiver diferente
SOURCE_SHEET = "Preços e Condições"        # ajuste apenas se o nome da aba estiver diferente
APP_TITLE = "Cockpit Papel"

# =========================================================
# CSS / LAYOUT
# =========================================================
st.markdown("""
<style>
/* sobe conteúdo e usa melhor o espaço */
.block-container {
    padding-top: 0.30rem !important;
    padding-bottom: 0.60rem !important;
    max-width: 100% !important;
}

/* título menor e compacto */
.custom-title {
    font-size: 1.35rem;
    font-weight: 700;
    margin-top: 0rem;
    margin-bottom: 0.40rem;
    line-height: 1.02;
    color: #1F2937;
}

/* subtítulos */
.section-title {
    font-size: 1.00rem;
    font-weight: 700;
    margin-top: 0.45rem;
    margin-bottom: 0.35rem;
    color: #1F2937;
}

/* cards KPI */
.kpi-card {
    background: linear-gradient(180deg, #FFFFFF 0%, #F7F9FC 100%);
    border: 1px solid #E5E7EB;
    border-radius: 14px;
    padding: 10px 12px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    min-height: 78px;
}

.kpi-label {
    font-size: 0.72rem;
    color: #6B7280;
    margin-bottom: 4px;
    font-weight: 600;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}

.kpi-value {
    font-size: 0.92rem;
    color: #111827;
    font-weight: 700;
    line-height: 1.15;
    word-break: break-word;
}

/* sidebar compacta */
[data-testid="stSidebar"] .block-container {
    padding-top: 0.75rem !important;
    padding-bottom: 0.65rem !important;
}

/* menos espaços internos */
div[data-testid="stVerticalBlock"] > div {
    gap: 0.35rem !important;
}

/* botões preenchendo a largura */
div[data-testid="stDownloadButton"] > button {
    width: 100%;
}
</style>
""", unsafe_allow_html=True)

st.markdown(f"<div class='custom-title'>{APP_TITLE}</div>", unsafe_allow_html=True)

# =========================================================
# HELPERS
# =========================================================
def normalize_text(value):
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("utf-8")
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def parse_number(value):
    """
    Converte valores em formatos variados para float.
    Exemplos:
    - 1.234,56
    - 1234,56
    - 1,234.56
    - R$ 1.234,56
    """
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()

    if s == "" or s.lower() in {"nan", "none", "-", "--"}:
        return None

    s = s.replace("R$", "").replace("$", "").replace("€", "")
    s = s.replace("%", "")
    s = s.replace(" ", "")
    s = re.sub(r"[^0-9,.\-]", "", s)

    if s in {"", "-", ".", ","}:
        return None

    if "," in s and "." in s:
        # padrão BR: 1.234,56
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            # padrão US: 1,234.56
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "")
        s = s.replace(",", ".")
    else:
        if s.count(".") > 1:
            parts = s.split(".")
            s = "".join(parts[:-1]) + "." + parts[-1]

    try:
        return float(s)
    except Exception:
        return None


def series_to_numeric(series):
    return series.apply(parse_number)


def format_br_number(value, decimals=2):
    """Retorna no padrão brasileiro: 1.234,56"""
    if pd.isna(value):
        return ""
    try:
        n = float(value)
    except Exception:
        return str(value)
    s = f"{n:,.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def format_no_decimal(value):
    """Sem casas decimais."""
    if pd.isna(value):
        return ""
    try:
        return str(int(round(float(value))))
    except Exception:
        return str(value)


def safe_display_string(value, numeric_no_decimal=False):
    """
    Converte qualquer valor para string segura para uso no multiselect.
    Isso evita o erro de proto.options no Streamlit.
    """
    if pd.isna(value):
        return ""

    if numeric_no_decimal:
        return format_no_decimal(value)

    # datas
    if isinstance(value, pd.Timestamp):
        if pd.isna(value):
            return ""
        return value.strftime("%d/%m/%Y")

    return str(value)


def detect_header_row(raw_df):
    """
    Procura a linha do cabeçalho real.
    """
    targets = ["impress type", "supplier", "current price"]

    best_row = None
    best_score = -1
    max_rows = min(len(raw_df), 50)

    for i in range(max_rows):
        row_values = [normalize_text(v) for v in raw_df.iloc[i].tolist()]
        row_text = " | ".join(row_values)

        score = sum(1 for t in targets if t in row_text)
        if score > best_score:
            best_score = score
            best_row = i

    return best_row if best_score >= 2 else None


def find_column(columns, aliases):
    normalized_map = {col: normalize_text(col) for col in columns}

    for alias in aliases:
        alias_norm = normalize_text(alias)

        for col, col_norm in normalized_map.items():
            if col_norm == alias_norm:
                return col

        for col, col_norm in normalized_map.items():
            if alias_norm in col_norm:
                return col

    return None


def build_canonical_dataframe(df):
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

    if "Última Atualização de Preço" in df2.columns:
        df2["Última Atualização de Preço"] = pd.to_datetime(
            df2["Última Atualização de Preço"],
            errors="coerce",
            dayfirst=True
        )

    # preenchimentos mínimos
    if "TCO (R$/KG)" not in df2.columns and "Current Price" in df2.columns:
        df2["TCO (R$/KG)"] = df2["Current Price"]

    if "P.Value (R$/KG)" not in df2.columns and "TCO (R$/KG)" in df2.columns:
        df2["P.Value (R$/KG)"] = df2["TCO (R$/KG)"]

    if "TCO (R$/M2)" not in df2.columns and {"TCO (R$/KG)", "g/m2"}.issubset(df2.columns):
        df2["TCO (R$/M2)"] = df2["TCO (R$/KG)"] * (df2["g/m2"] / 1000.0)

    if "P.Value (R$/M2)" not in df2.columns and {"P.Value (R$/KG)", "g/m2"}.issubset(df2.columns):
        df2["P.Value (R$/M2)"] = df2["P.Value (R$/KG)"] * (df2["g/m2"] / 1000.0)

    # limpeza mínima
    essential = []
    if "Impress Type" in df2.columns:
        essential.append("Impress Type")
    if "Supplier" in df2.columns:
        essential.append("Supplier")

    if essential:
        df2 = df2.dropna(subset=essential)

    if "Current Price" in df2.columns:
        df2 = df2[df2["Current Price"].notna()]
    elif "TCO (R$/KG)" in df2.columns:
        df2 = df2[df2["TCO (R$/KG)"].notna()]

    sort_cols = [c for c in ["Impress Type", "Width (mm)", "g/m2", "Supplier", "Lot (ton)"] if c in df2.columns]
    if sort_cols:
        df2 = df2.sort_values(sort_cols, kind="stable")

    return df2


def create_safe_multiselect(df_in, column_name, label, numeric_no_decimal=False):
    """
    Cria multiselect com opções seguras como string, evitando TypeError no Streamlit.
    """
    if column_name not in df_in.columns:
        return df_in

    source = df_in[column_name].dropna().copy()
    if source.empty:
        return df_in

    # Tabela de opções: original x texto exibido
    opt_df = pd.DataFrame({"original": source})
    opt_df["display"] = opt_df["original"].apply(
        lambda x: safe_display_string(x, numeric_no_decimal=numeric_no_decimal)
    )

    # Remove duplicados de display preservando 1º original
    opt_df = opt_df.drop_duplicates(subset=["display"], keep="first")

    # Ordenação amigável
    if numeric_no_decimal:
        opt_df["sort_key"] = opt_df["original"].apply(lambda x: parse_number(x) if parse_number(x) is not None else 10**18)
        opt_df = opt_df.sort_values(["sort_key", "display"], kind="stable")
    else:
        opt_df["sort_key"] = opt_df["display"].astype(str)
        opt_df = opt_df.sort_values("sort_key", kind="stable")

    display_options = opt_df["display"].tolist()

    selected_display = st.sidebar.multiselect(
        label,
        options=display_options,
        default=[]
    )

    if not selected_display:
        return df_in

    selected_original = opt_df.loc[opt_df["display"].isin(selected_display), "original"].tolist()
    return df_in[df_in[column_name].isin(selected_original)]


def safe_min(series):
    s = pd.to_numeric(series, errors="coerce").dropna()
    return s.min() if not s.empty else None


def safe_best_row(df_in, metric_col):
    if metric_col not in df_in.columns:
        return None
    tmp = df_in.copy()
    tmp[metric_col] = pd.to_numeric(tmp[metric_col], errors="coerce")
    tmp = tmp.dropna(subset=[metric_col])
    if tmp.empty:
        return None
    idx = tmp[metric_col].idxmin()
    return tmp.loc[idx]


def kpi_card(label, value):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def to_excel_bytes(df_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Cockpit Filtrado")
    output.seek(0)
    return output.getvalue()

# =========================================================
# LOAD
# =========================================================
@st.cache_data
def load_data():
    raw = pd.read_excel(EXCEL_FILE, sheet_name=SOURCE_SHEET, header=None)

    header_row = detect_header_row(raw)
    if header_row is None:
        raise ValueError(
            "Não foi possível identificar automaticamente a linha de cabeçalho da aba 'Preços e Condições'."
        )

    df = pd.read_excel(EXCEL_FILE, sheet_name=SOURCE_SHEET, header=header_row)
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]
    df.columns = [re.sub(r"\s+", " ", c) for c in df.columns]
    df = df.dropna(how="all")

    return build_canonical_dataframe(df)


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
# SIDEBAR = FILTROS
# =========================================================
st.sidebar.header("Filtros")

filtered = df.copy()

# Ordem dos slicers
filtered = create_safe_multiselect(filtered, "Impress Type", "Impress Type", numeric_no_decimal=False)
filtered = create_safe_multiselect(filtered, "Width (mm)", "Width (mm)", numeric_no_decimal=True)
filtered = create_safe_multiselect(filtered, "g/m2", "g/m2", numeric_no_decimal=True)
filtered = create_safe_multiselect(filtered, "Supplier", "Supplier", numeric_no_decimal=False)
filtered = create_safe_multiselect(filtered, "Currency", "Currency", numeric_no_decimal=False)
filtered = create_safe_multiselect(filtered, "Lot (ton)", "Lot (ton)", numeric_no_decimal=True)

# -----------------------------------------
# BLOQUEIO SEM FILTRO
# -----------------------------------------
if filtered.shape[0] == df.shapest.info("Selecione pelo menos um filtro para visualizar os dados.")
    st.stop()

if filtered.empty:
    st.warning("Nenhum registro encontrado com os filtros selecionados.")
    st.stop()

# =========================================================
# KPI
# =========================================================
best_pv_kg = safe_min(filtered["P.Value (R$/KG)"]) if "P.Value (R$/KG)" in filtered.columns else None
best_pv_m2 = safe_min(filtered["P.Value (R$/M2)"]) if "P.Value (R$/M2)" in filtered.columns else None
best_row = safe_best_row(filtered, "P.Value (R$/KG)") if "P.Value (R$/KG)" in filtered.columns else None

best_supplier = best_row["Supplier"] if best_row is not None and "Supplier" in filtered.columns else "N/A"

if "Impress Type" in filtered.columns:
    unique_impress = filtered["Impress Type"].dropna().unique().tolist()
    if len(unique_impress) == 1:
        impress_value = str(unique_impress[0])
    else:
        impress_value = f"{len(unique_impress)} tipos"
else:
    impress_value = "N/A"

k1, k2, k3, k4 = st.columns(4)

with k1:
    kpi_card("Impress Type", impress_value)

with k2:
    kpi_card(
        "Melhor P.Value (R$/KG)",
        format_br_number(best_pv_kg, 2) if best_pv_kg is not None else "N/A"
    )

with k3:
    kpi_card(
        "Melhor P.Value (R$/M2)",
        format_br_number(best_pv_m2, 2) if best_pv_m2 is not None else "N/A"
    )

with k4:
    kpi_card("Melhor Supplier", best_supplier if best_supplier else "N/A")

# =========================================================
# TABELA PRINCIPAL + EXPORTAÇÃO
# =========================================================
st.markdown("<div class='section-title'>Tabela principal</div>", unsafe_allow_html=True)

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
table_df_raw = filtered[display_cols].copy()

# renomeia apenas para exibição/exportação
table_df_raw = table_df_raw.rename(columns={
    "Última Atualização de Preço": "Último Preço"
})

# cópia formatada para exibir/exportar
table_df_display = table_df_raw.copy()

value_cols = [
    "Current Price",
    "TCO (R$/KG)",
    "TCO (R$/M2)",
    "P.Value (R$/KG)",
    "P.Value (R$/M2)",
]

no_decimal_cols = [
    "Width (mm)",
    "g/m2",
    "Paper bonus (t)",
    "Lot (ton)",
    "Working days",
]

for col in value_cols:
    if col in table_df_display.columns:
        table_df_display[col] = table_df_display[col].apply(lambda x: format_br_number(x, 2))

for col in no_decimal_cols:
    if col in table_df_display.columns:
        table_df_display[col] = table_df_display[col].apply(lambda x: "" if pd.isna(x) else format_no_decimal(x))

if "Último Preço" in table_df_display.columns:
    table_df_display["Último Preço"] = pd.to_datetime(
        table_df_display["Último Preço"],
        errors="coerce"
    ).dt.strftime("%d/%m/%Y")

# exportação
exp1, exp2, exp3 = st.columns([1.2, 1.2, 6])

with exp1:
    csv_bytes = table_df_display.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        label="Exportar CSV",
        data=csv_bytes,
        file_name="cockpit_filtrado.csv",
        mime="text/csv",
        use_container_width=True
    )

with exp2:
    excel_bytes = to_excel_bytes(table_df_display)
    st.download_button(
        label="Exportar Excel",
        data=excel_bytes,
        file_name="cockpit_filtrado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# alinhamento
all_cols = list(table_df_display.columns)
left_cols = [c for c in ["Supplier"] if c in table_df_display.columns]
right_cols = [c for c in value_cols if c in table_df_display.columns]
center_cols = [c for c in all_cols if c not in left_cols + right_cols]

styled_table = (
    table_df_display.style
    .set_properties(subset=center_cols, **{"text-align": "center"})
    .set_properties(subset=left_cols, **{"text-align": "left"})
    .set_properties(subset=right_cols, **{"text-align": "right"})
    .set_table_styles([
        {"selector": "th", "props": [("text-align", "center"), ("font-size", "12px")]},
        {"selector": "td", "props": [("font-size", "12px"), ("white-space", "nowrap")]},
    ])
)

st.dataframe(
    styled_table,
    use_container_width=True,
    hide_index=True,
    height=540,
    column_config={
        "Impress Type": st.column_config.TextColumn("Impress Type", width="medium"),
        "Width (mm)": st.column_config.TextColumn("Width (mm)", width="small"),
        "g/m2": st.column_config.TextColumn("g/m2", width="small"),
        "Supplier": st.column_config.TextColumn("Supplier", width="large"),
        "Currency": st.column_config.TextColumn("Currency", width="small"),
        "Current Price": st.column_config.TextColumn("Current Price", width="medium"),
        "Paper bonus (t)": st.column_config.TextColumn("Paper bonus (t)", width="small"),
        "Lot (ton)": st.column_config.TextColumn("Lot (ton)", width="small"),
        "TCO (R$/KG)": st.column_config.TextColumn("TCO (R$/KG)", width="medium"),
        "TCO (R$/M2)": st.column_config.TextColumn("TCO (R$/M2)", width="medium"),
        "Payment Terms": st.column_config.TextColumn("Payment Terms", width="medium"),
        "Working days": st.column_config.TextColumn("Working days", width="small"),
        "P.Value (R$/KG)": st.column_config.TextColumn("P.Value (R$/KG)", width="medium"),
        "P.Value (R$/M2)": st.column_config.TextColumn("P.Value (R$/M2)", width="medium"),
        "Último Preço": st.column_config.TextColumn("Último Preço", width="medium"),
    }
)

# =========================================================
# GRÁFICO TCO MÉDIO (ABAIXO DA TABELA)
# =========================================================
st.markdown("<div class='section-title'>TCO médio por Supplier</div>", unsafe_allow_html=True)

if {"Supplier", "TCO (R$/KG)"}.issubset(filtered.columns):
    chart_df = filtered.copy()
    chart_df["TCO (R$/KG)"] = pd.to_numeric(chart_df["TCO (R$/KG)"], errors="coerce")
    chart_df = chart_df.dropna(subset=["Supplier", "TCO (R$/KG)"])

    if not chart_df.empty:
        chart_base = (
            chart_df.groupby("Supplier", as_index=False)["TCO (R$/KG)"]
            .mean()
            .sort_values("TCO (R$/KG)", ascending=True)
        )

        chart_base["TCO_label"] = chart_base["TCO (R$/KG)"].apply(lambda x: format_br_number(x, 2))

        fig_bar = px.bar(
            chart_base,
            x="Supplier",
            y="TCO (R$/KG)",
            color="TCO (R$/KG)",
            color_continuous_scale="Blues",
            text="TCO_label"
        )

        fig_bar.update_traces(textposition="outside")
        fig_bar.update_layout(
            showlegend=False,
            height=390,
            margin=dict(t=20, r=20, l=20, b=20),
            xaxis_title="Supplier",
            yaxis_title="TCO (R$/KG)",
            plot_bgcolor="white",
            paper_bgcolor="white"
        )

        fig_bar.update_xaxes(showgrid=False)
        fig_bar.update_yaxes(showgrid=True, gridcolor="#E5E7EB")

        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("Sem dados válidos para exibir o gráfico.")
else:
    st.info("As colunas necessárias para o gráfico não estão disponíveis.")

# =========================================================
# RODAPÉ
# =========================================================
st.caption(
    "O dashboard considera a aba 'Preços e Condições' como base, "
    "usa apenas 'Current Price' como preço principal e ignora colunas históricas mensais."
)
