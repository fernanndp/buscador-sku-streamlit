import io
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Set

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz


st.set_page_config(page_title="Buscador de SKU", page_icon="🔎", layout="wide")


REFERENCE_EAN_ALIASES = [
    "EAN 13 - UND", "EAN", "Código de Barras", "Codigo de Barras",
    "GTIN", "Código Barras", "Codigo Barras"
]
REFERENCE_DESC_ALIASES = [
    "Produtos", "Produto", "Descrição", "Descricao", "Descrição Produto", "Descricao Produto"
]
TARGET_EAN_ALIASES = [
    "Codigo Barras", "Código Barras", "Código de Barras", "Codigo de Barras",
    "GTIN", "EAN", "CODIGO_BARRAS"
]
TARGET_DESC_ALIASES = [
    "Descripcion", "Descrição", "Descricao", "Produto", "Produtos", "TOP_DESCRIPCION", "MAX_DESCRIPCION"
]


@dataclass
class DictionaryConfig:
    rules: pd.DataFrame
    stopwords: Set[str]
    category_noise: Set[str]
    brand_hints: Set[str]


def strip_accents(text: str) -> str:
    text = unicodedata.normalize("NFKD", str(text))
    return "".join(char for char in text if not unicodedata.combining(char))


def normalize_header(text: str) -> str:
    text = strip_accents(text).lower().strip()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_vocab_term(value: str) -> str:
    value = strip_accents(str(value)).lower().strip()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def normalize_sheet_name(name: str) -> str:
    return normalize_header(name).replace(" ", "")


def detect_column(columns: Iterable[str], aliases: List[str]) -> Optional[str]:
    normalized_map = {normalize_header(col): col for col in columns}
    for alias in aliases:
        key = normalize_header(alias)
        if key in normalized_map:
            return normalized_map[key]
    return None


def load_dictionary_template_bytes() -> bytes:
    template_path = Path(__file__).parent / "dicionario_template_sku.xlsx"
    if not template_path.exists():
        raise FileNotFoundError("Arquivo 'dicionario_template_sku.xlsx' não encontrado na pasta do projeto.")
    return template_path.read_bytes()


def read_uploaded_table(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    content = uploaded_file.getvalue()

    if name.endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(content), engine="openpyxl")

    if name.endswith(".xls"):
        return pd.read_excel(io.BytesIO(content), engine="xlrd")

    if name.endswith(".csv"):
        attempts = [
            {"encoding": "utf-8-sig", "sep": None, "engine": "python"},
            {"encoding": "utf-8", "sep": None, "engine": "python"},
            {"encoding": "utf-16", "sep": None, "engine": "python"},
            {"encoding": "latin-1", "sep": None, "engine": "python"},
            {"encoding": "utf-16", "sep": ";"},
            {"encoding": "latin-1", "sep": ";"},
            {"encoding": "utf-8-sig", "sep": ";"},
        ]
        last_error = None
        for kwargs in attempts:
            try:
                return pd.read_csv(io.BytesIO(content), **kwargs)
            except Exception as exc:
                last_error = exc
        raise ValueError(f"Não consegui ler o CSV. Erro: {last_error}")

    raise ValueError("Formato não suportado. Use .xls, .xlsx ou .csv")


def load_excel_sheets(uploaded_file) -> dict[str, pd.DataFrame]:
    name = uploaded_file.name.lower()
    content = uploaded_file.getvalue()

    if name.endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(content), sheet_name=None, engine="openpyxl")

    if name.endswith(".xls"):
        return pd.read_excel(io.BytesIO(content), sheet_name=None, engine="xlrd")

    raise ValueError("Para usar abas de configuração, envie o dicionário em .xlsx ou .xls")


def normalize_rules_frame(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {}
    for col in df.columns:
        col_norm = normalize_header(col)
        if col_norm in {"padrao", "padrão", "regex", "regra"}:
            rename_map[col] = "Padrao"
        elif col_norm in {"substituto", "substituicao", "substituição", "troca", "valor"}:
            rename_map[col] = "Substituto"

    df = df.rename(columns=rename_map)

    if "Padrao" not in df.columns or "Substituto" not in df.columns:
        if len(df.columns) >= 2:
            df = df.iloc[:, :2].copy()
            df.columns = ["Padrao", "Substituto"]
        else:
            raise ValueError("A aba de regras precisa ter pelo menos 2 colunas: Padrao e Substituto.")

    df = df[["Padrao", "Substituto"]].copy()
    df["Padrao"] = df["Padrao"].fillna("").astype(str)
    df["Substituto"] = df["Substituto"].fillna("").astype(str)
    df = df[(df["Padrao"].str.strip() != "")]
    return df.drop_duplicates().reset_index(drop=True)


def first_text_column(df: pd.DataFrame) -> str:
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            return col
    return df.columns[0]


def normalize_term_set(df: pd.DataFrame) -> Set[str]:
    if df.empty:
        return set()

    chosen_col = None
    for col in df.columns:
        col_norm = normalize_header(col)
        if col_norm in {"palavra", "termo", "token", "valor", "item"}:
            chosen_col = col
            break

    if chosen_col is None:
        chosen_col = first_text_column(df)

    values = {
        normalize_vocab_term(value)
        for value in df[chosen_col].fillna("").astype(str)
        if normalize_vocab_term(value)
    }
    return values


def read_dictionary(uploaded_file) -> DictionaryConfig:
    if uploaded_file is None:
        raise ValueError("Envie o dicionário. O app não usa nenhuma regra embutida no código.")

    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        rules_df = normalize_rules_frame(read_uploaded_table(uploaded_file))
        return DictionaryConfig(
            rules=rules_df,
            stopwords=set(),
            category_noise=set(),
            brand_hints=set(),
        )

    workbook = load_excel_sheets(uploaded_file)
    normalized_sheet_map = {normalize_sheet_name(sheet_name): sheet_name for sheet_name in workbook.keys()}

    def get_sheet(*aliases: str) -> Optional[pd.DataFrame]:
        for alias in aliases:
            normalized_alias = normalize_sheet_name(alias)
            real_name = normalized_sheet_map.get(normalized_alias)
            if real_name:
                return workbook[real_name]
        return None

    rules_sheet = get_sheet("Regras", "Dicionario", "Dictionary", "Padroes")
    stopwords_sheet = get_sheet("Stopwords", "StopWords", "PalavrasIgnoradas")
    category_noise_sheet = get_sheet("CategoryNoise", "Noise", "RuidoCategoria", "Ruido", "CategoriasIgnoradas")
    brand_hints_sheet = get_sheet("BrandHints", "Marcas", "Brand", "Hints")

    if rules_sheet is None:
        raise ValueError("A aba 'Regras' é obrigatória no dicionário.")

    rules_df = normalize_rules_frame(rules_sheet)
    stopwords = normalize_term_set(stopwords_sheet) if stopwords_sheet is not None else set()
    category_noise = normalize_term_set(category_noise_sheet) if category_noise_sheet is not None else set()
    brand_hints = normalize_term_set(brand_hints_sheet) if brand_hints_sheet is not None else set()

    return DictionaryConfig(
        rules=rules_df,
        stopwords=stopwords,
        category_noise=category_noise,
        brand_hints=brand_hints,
    )


def normalize_ean(value) -> str:
    digits = re.sub(r"\D", "", "" if pd.isna(value) else str(value))
    return digits.strip()


def normalize_description(text: str, rules: pd.DataFrame) -> str:
    text = "" if pd.isna(text) else str(text)
    text = strip_accents(text).lower().replace(",", ".")
    text = text.replace("/", " / ")

    for pattern, replacement in rules[["Padrao", "Substituto"]].itertuples(index=False):
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)

    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def tokenize(text: str, rules: pd.DataFrame, stopwords: Set[str]) -> List[str]:
    normalized = normalize_description(text, rules)
    tokens = normalized.split()
    return [token for token in tokens if token not in stopwords]


def extract_quantity(text: str, rules: pd.DataFrame) -> str:
    normalized = normalize_description(text, rules)

    pack_match = re.search(r"\b(\d+x\d+(?:\.\d+)?(?:g|kg|ml|l))\b", normalized)
    if pack_match:
        return pack_match.group(1)

    simple_match = re.search(r"\b(\d+(?:\.\d+)?(?:g|kg|ml|l))\b", normalized)
    return simple_match.group(1) if simple_match else ""


def signature(text: str, rules: pd.DataFrame, stopwords: Set[str]) -> str:
    return " ".join(sorted(set(tokenize(text, rules, stopwords))))


def ratio_intersection(
    reference_text: str,
    target_text: str,
    rules: pd.DataFrame,
    stopwords: Set[str],
    category_noise: Set[str],
) -> float:
    left = {token for token in tokenize(reference_text, rules, stopwords) if token not in category_noise}
    right = {token for token in tokenize(target_text, rules, stopwords) if token not in category_noise}

    if not left:
        return 0.0

    return len(left & right) / len(left)


def similarity_score(
    reference_text: str,
    target_text: str,
    rules: pd.DataFrame,
    stopwords: Set[str],
    category_noise: Set[str],
) -> float:
    normalized_ref = normalize_description(reference_text, rules)
    normalized_tgt = normalize_description(target_text, rules)

    signature_ref = signature(reference_text, rules, stopwords)
    signature_tgt = signature(target_text, rules, stopwords)

    qty_ref = extract_quantity(reference_text, rules)
    qty_tgt = extract_quantity(target_text, rules)

    score = (
        0.45 * fuzz.token_set_ratio(normalized_ref, normalized_tgt)
        + 0.25 * fuzz.partial_ratio(signature_ref, signature_tgt)
        + 0.20 * (ratio_intersection(reference_text, target_text, rules, stopwords, category_noise) * 100)
        + 0.10 * (100 if qty_ref and qty_ref == qty_tgt else 0)
    )

    return round(score / 100, 4)


def prepare_frame(df: pd.DataFrame, ean_col: str, desc_col: str, rules: pd.DataFrame, stopwords: Set[str]) -> pd.DataFrame:
    prepared = df.copy()
    prepared["_ean_norm"] = prepared[ean_col].apply(normalize_ean)
    prepared["_desc_raw"] = prepared[desc_col].fillna("").astype(str)
    prepared["_desc_norm"] = prepared["_desc_raw"].apply(lambda value: normalize_description(value, rules))
    prepared["_tokens"] = prepared["_desc_raw"].apply(lambda value: tokenize(value, rules, stopwords))
    prepared["_token_set"] = prepared["_tokens"].apply(set)
    prepared["_signature"] = prepared["_desc_raw"].apply(lambda value: signature(value, rules, stopwords))
    prepared["_qty"] = prepared["_desc_raw"].apply(lambda value: extract_quantity(value, rules))
    return prepared


def prefilter_candidates(reference_row: pd.Series, target_df: pd.DataFrame, brand_hints: Set[str]) -> pd.DataFrame:
    candidates = target_df

    qty = reference_row["_qty"]
    if qty:
        same_qty = candidates[candidates["_qty"] == qty]
        if not same_qty.empty:
            candidates = same_qty

    brand_tokens = reference_row["_token_set"] & brand_hints
    if brand_tokens:
        brand_mask = candidates["_token_set"].apply(lambda token_set: bool(token_set & brand_tokens))
        filtered = candidates[brand_mask]
        if not filtered.empty:
            candidates = filtered

    if len(candidates) > 800:
        ref_tokens = reference_row["_token_set"]
        candidates = candidates.assign(
            _prefilter_overlap=candidates["_token_set"].apply(lambda token_set: len(ref_tokens & token_set))
        )
        candidates = candidates.sort_values("_prefilter_overlap", ascending=False).head(150).drop(columns="_prefilter_overlap")

    return candidates


def classify_status(ean_found: bool, score: float, desc_threshold: float, suggestion_threshold: float) -> str:
    if ean_found and score >= desc_threshold:
        return "OK"

    if ean_found and score < desc_threshold:
        return "EAN encontrado / descrição revisar"

    if (not ean_found) and score >= suggestion_threshold:
        return "EAN não encontrado / descrição semelhante"

    return "EAN não encontrado"


def run_comparison(
    reference_df: pd.DataFrame,
    target_df: pd.DataFrame,
    reference_ean_col: str,
    reference_desc_col: str,
    target_ean_col: str,
    target_desc_col: str,
    rules: pd.DataFrame,
    stopwords: Set[str],
    category_noise: Set[str],
    brand_hints: Set[str],
    desc_threshold: float,
    suggestion_threshold: float,
) -> pd.DataFrame:
    reference = prepare_frame(reference_df, reference_ean_col, reference_desc_col, rules, stopwords)
    target = prepare_frame(target_df, target_ean_col, target_desc_col, rules, stopwords)

    target_ean_map = {
        ean: group.copy()
        for ean, group in target.groupby("_ean_norm", dropna=False)
        if str(ean).strip() != ""
    }

    output_rows = []

    for _, ref_row in reference.iterrows():
        ref_ean = ref_row["_ean_norm"]
        ean_candidates = target_ean_map.get(ref_ean)

        if ean_candidates is not None and not ean_candidates.empty:
            candidates = ean_candidates
            ean_found = True
        else:
            candidates = prefilter_candidates(ref_row, target, brand_hints)
            ean_found = False

        if candidates.empty:
            best_score = 0.0
            best_target_ean = ""
            best_target_desc = ""
        else:
            scored = candidates[[target_ean_col, target_desc_col, "_desc_raw"]].copy()
            scored["similaridade"] = candidates["_desc_raw"].apply(
                lambda value: similarity_score(ref_row["_desc_raw"], value, rules, stopwords, category_noise)
            )

            best_idx = scored["similaridade"].idxmax()
            best_score = float(scored.loc[best_idx, "similaridade"])
            best_target_ean = scored.loc[best_idx, target_ean_col]
            best_target_desc = scored.loc[best_idx, target_desc_col]

        output_rows.append(
            {
                "EAN_ref": ref_row[reference_ean_col],
                "Descricao_ref": ref_row[reference_desc_col],
                "EAN_ref_norm": ref_ean,
                "EAN_encontrado": "Sim" if ean_found else "Não",
                "Descricao_target": best_target_desc,
                "EAN_target": best_target_ean,
                "EAN_target_norm": normalize_ean(best_target_ean),
                "Similaridade": round(best_score, 4),
                "Descricao_ok": "Sim" if best_score >= desc_threshold else "Não",
                "Status": classify_status(ean_found, best_score, desc_threshold, suggestion_threshold),
                "Descricao_ref_normalizada": ref_row["_desc_norm"],
                "Descricao_target_normalizada": normalize_description(best_target_desc, rules),
            }
        )

    return pd.DataFrame(output_rows)


def build_excel_output(
    result_df: pd.DataFrame,
    effective_dictionary: pd.DataFrame,
    effective_stopwords: Set[str],
    effective_category_noise: Set[str],
    effective_brand_hints: Set[str],
    reference_df: pd.DataFrame,
    target_df: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()

    resumo = pd.DataFrame(
        [
            ["Total SKUs referência", len(result_df)],
            ["EAN encontrados", int((result_df["EAN_encontrado"] == "Sim").sum())],
            ["EAN faltando", int((result_df["EAN_encontrado"] == "Não").sum())],
            ["Descrições aprovadas", int((result_df["Descricao_ok"] == "Sim").sum())],
            ["Descrições para revisão", int((result_df["Descricao_ok"] == "Não").sum())],
            ["Status OK", int((result_df["Status"] == "OK").sum())],
        ],
        columns=["Métrica", "Valor"],
    )

    stopwords_df = pd.DataFrame(sorted(effective_stopwords), columns=["Palavra"])
    category_noise_df = pd.DataFrame(sorted(effective_category_noise), columns=["Palavra"])
    brand_hints_df = pd.DataFrame(sorted(effective_brand_hints), columns=["Palavra"])

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
        result_df.to_excel(writer, sheet_name="Comparacao", index=False)
        result_df[result_df["EAN_encontrado"] == "Não"].to_excel(writer, sheet_name="EAN_Faltando", index=False)
        result_df[result_df["Descricao_ok"] == "Não"].to_excel(writer, sheet_name="Descricao_Revisar", index=False)
        effective_dictionary.to_excel(writer, sheet_name="Regras_Efetivas", index=False)
        stopwords_df.to_excel(writer, sheet_name="Stopwords_Efetivas", index=False)
        category_noise_df.to_excel(writer, sheet_name="Noise_Efetivo", index=False)
        brand_hints_df.to_excel(writer, sheet_name="BrandHints_Efetivo", index=False)
        reference_df.head(200).to_excel(writer, sheet_name="Preview_Referencia", index=False)
        target_df.head(200).to_excel(writer, sheet_name="Preview_Busca", index=False)

        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column_cells in worksheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter
                for cell in column_cells:
                    try:
                        cell_value = "" if cell.value is None else str(cell.value)
                        max_length = max(max_length, len(cell_value))
                    except Exception:
                        pass
                worksheet.column_dimensions[column_letter].width = min(max_length + 2, 45)

    return output.getvalue()


def show_metrics(result_df: pd.DataFrame) -> None:
    total = len(result_df)
    eans_found = int((result_df["EAN_encontrado"] == "Sim").sum())
    eans_missing = int((result_df["EAN_encontrado"] == "Não").sum())
    descriptions_ok = int((result_df["Descricao_ok"] == "Sim").sum())

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("SKUs referência", total)
    col2.metric("EAN encontrados", eans_found)
    col3.metric("EAN faltando", eans_missing)
    col4.metric("Descrições OK", descriptions_ok)


def main() -> None:
    st.title("🔎 Buscador de SKU")
    st.caption("Compare uma planilha de referência contra outra planilha de busca, validando EAN e similaridade de descrição.")

    with st.expander("Como funciona", expanded=True):
        st.markdown(
            """
            1. Envie a planilha base com os SKUs de referência.  
            2. Envie a planilha onde esses SKUs devem ser buscados.  
            3. Envie o dicionário de equivalências.  
            4. O app verifica:
               - se o EAN da planilha base aparece na planilha de busca;
               - se a descrição da base é semelhante à melhor descrição encontrada.
            """
        )

    st.subheader("📘 Dicionário de equivalências")
    st.caption("Baixe o template e envie o dicionário preenchido.")

    try:
        template_bytes = load_dictionary_template_bytes()
        st.download_button(
            label="📥 Baixar template do dicionário",
            data=template_bytes,
            file_name="dicionario_template_sku.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=False,
        )
    except FileNotFoundError as exc:
        st.warning(str(exc))

    left, right = st.columns(2)
    with left:
        reference_file = st.file_uploader(
            "Planilha 1 — referência de produtos",
            type=["xls", "xlsx", "csv"],
            key="reference_file",
        )
    with right:
        target_file = st.file_uploader(
            "Planilha 2 — planilha onde será feita a busca",
            type=["xls", "xlsx", "csv"],
            key="target_file",
        )

    dictionary_file = st.file_uploader(
        "Dicionário de equivalências/configuração",
        type=["xls", "xlsx", "csv"],
        key="dictionary_file",
    )

    if not reference_file or not target_file:
        st.info("Envie as duas planilhas para continuar.")
        return

    try:
        reference_df = read_uploaded_table(reference_file)
        target_df = read_uploaded_table(target_file)
        dictionary_config = read_dictionary(dictionary_file)
        effective_dictionary = dictionary_config.rules.copy()
        effective_stopwords = set(dictionary_config.stopwords)
        effective_category_noise = set(dictionary_config.category_noise)
        effective_brand_hints = set(dictionary_config.brand_hints)
    except Exception as exc:
        st.error(f"Erro ao ler arquivos: {exc}")
        return

    ref_ean_default = detect_column(reference_df.columns, REFERENCE_EAN_ALIASES)
    ref_desc_default = detect_column(reference_df.columns, REFERENCE_DESC_ALIASES)
    tgt_ean_default = detect_column(target_df.columns, TARGET_EAN_ALIASES)
    tgt_desc_default = detect_column(target_df.columns, TARGET_DESC_ALIASES)

    st.subheader("Mapeamento das colunas")

    cfg1, cfg2 = st.columns(2)
    with cfg1:
        reference_ean_col = st.selectbox(
            "Coluna de EAN — planilha 1",
            options=list(reference_df.columns),
            index=list(reference_df.columns).index(ref_ean_default) if ref_ean_default in reference_df.columns else 0,
        )
        reference_desc_col = st.selectbox(
            "Coluna de descrição — planilha 1",
            options=list(reference_df.columns),
            index=list(reference_df.columns).index(ref_desc_default) if ref_desc_default in reference_df.columns else 0,
        )

    with cfg2:
        target_ean_col = st.selectbox(
            "Coluna de EAN — planilha 2",
            options=list(target_df.columns),
            index=list(target_df.columns).index(tgt_ean_default) if tgt_ean_default in target_df.columns else 0,
        )
        target_desc_col = st.selectbox(
            "Coluna de descrição — planilha 2",
            options=list(target_df.columns),
            index=list(target_df.columns).index(tgt_desc_default) if tgt_desc_default in target_df.columns else 0,
        )

    params1, params2 = st.columns(2)
    with params1:
        desc_threshold = st.slider(
            "Limite para aprovar descrição",
            min_value=0.50,
            max_value=1.00,
            value=0.75,
            step=0.01,
            help="Quanto maior, mais rigorosa fica a validação de descrição.",
        )
    with params2:
        suggestion_threshold = st.slider(
            "Limite para marcar sugestão sem EAN",
            min_value=0.50,
            max_value=1.00,
            value=0.72,
            step=0.01,
            help="Usado quando o EAN não foi encontrado, mas a descrição parece muito semelhante.",
        )

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "Prévia planilha 1",
        "Prévia planilha 2",
        "Regras efetivas",
        "Stopwords",
        "CategoryNoise",
        "BrandHints",
    ])

    with tab1:
        st.dataframe(reference_df[[reference_ean_col, reference_desc_col]].head(20), use_container_width=True)

    with tab2:
        st.dataframe(target_df[[target_ean_col, target_desc_col]].head(20), use_container_width=True)

    with tab3:
        st.dataframe(effective_dictionary, use_container_width=True, hide_index=True)

    with tab4:
        st.dataframe(pd.DataFrame(sorted(effective_stopwords), columns=["Palavra"]), use_container_width=True, hide_index=True)

    with tab5:
        st.dataframe(pd.DataFrame(sorted(effective_category_noise), columns=["Palavra"]), use_container_width=True, hide_index=True)

    with tab6:
        st.dataframe(pd.DataFrame(sorted(effective_brand_hints), columns=["Palavra"]), use_container_width=True, hide_index=True)

    if st.button("Comparar SKUs", type="primary", use_container_width=True):
        with st.spinner("Comparando EANs e descrições..."):
            result_df = run_comparison(
                reference_df=reference_df,
                target_df=target_df,
                reference_ean_col=reference_ean_col,
                reference_desc_col=reference_desc_col,
                target_ean_col=target_ean_col,
                target_desc_col=target_desc_col,
                rules=effective_dictionary,
                stopwords=effective_stopwords,
                category_noise=effective_category_noise,
                brand_hints=effective_brand_hints,
                desc_threshold=desc_threshold,
                suggestion_threshold=suggestion_threshold,
            )

        show_metrics(result_df)

        result_tabs = st.tabs(["Resultado completo", "EAN faltando", "Descrição revisar"])

        with result_tabs[0]:
            st.dataframe(result_df, use_container_width=True, hide_index=True)

        with result_tabs[1]:
            st.dataframe(
                result_df[result_df["EAN_encontrado"] == "Não"],
                use_container_width=True,
                hide_index=True,
            )

        with result_tabs[2]:
            st.dataframe(
                result_df[result_df["Descricao_ok"] == "Não"],
                use_container_width=True,
                hide_index=True,
            )

        excel_bytes = build_excel_output(
            result_df=result_df,
            effective_dictionary=effective_dictionary,
            effective_stopwords=effective_stopwords,
            effective_category_noise=effective_category_noise,
            effective_brand_hints=effective_brand_hints,
            reference_df=reference_df,
            target_df=target_df,
        )

        st.download_button(
            label="Baixar resultado em Excel",
            data=excel_bytes,
            file_name="resultado_buscador_sku.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()