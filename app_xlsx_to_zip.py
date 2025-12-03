# ==============================================================================
#  app_xlsx_to_zip.py — Versão MULTI-XLSX + Fix Total do Erro PyArrow
#  Inclui exibição segura via make_safe_display_df()
# ==============================================================================

import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import re
from typing import Dict, List, Tuple

st.set_page_config(page_title="XLSX → CSV (ZIP) | Múltiplos Arquivos", layout="wide")

# ==============================================================================
#  FIX DEFINITIVO DO WARNING PYARROW
# ==============================================================================

def make_safe_display_df(df: pd.DataFrame, max_rows: int = 200) -> pd.DataFrame:
    """
    Cria DataFrame seguro para exibição no Streamlit sem acionar erros PyArrow.
    Regras:
      - Trabalha em head(max_rows)
      - Para dtypes "object"/"string":
          * tenta converter vírgula decimal -> ponto e converter a numérico
          * SE >= 50% da coluna vira numérico => converte
          * SENÃO => força str (isso PREVINE o PyArrow error)
      - Não altera os dados originais, apenas a cópia exibida
    """
    if df is None or df.empty:
        return df.copy()

    out = df.head(max_rows).copy()

    for col in out.columns:
        s = out[col]

        # Se já é numérico → mantenha
        if pd.api.types.is_numeric_dtype(s):
            continue

        # Apenas para colunas de texto
        if pd.api.types.is_object_dtype(s) or pd.api.types.is_string_dtype(s):
            s2 = s.astype(str).str.strip()

            # Detecta se há "." e "," simultaneamente
            has_dot = s2.str.contains(r"\.", regex=True).any()
            has_comma = s2.str.contains(r",", regex=True).any()

            # Normaliza vírgula decimal / milhares
            if has_dot and has_comma:
                # remove separador de milhar ".", troca "," -> "."
                s2 = s2.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
            else:
                s2 = s2.str.replace(",", ".", regex=False)

            # tenta conversão
            num = pd.to_numeric(s2, errors="coerce")
            ratio = num.notna().sum() / len(num)

            if ratio >= 0.5:     # heurística
                out[col] = num
            else:
                out[col] = s.fillna("").astype(str).replace({"nan": "", "None": ""})
        else:
            # força para string se aparecer col numérica-misturada exótica
            try:
                out[col] = s.fillna("").astype(str)
            except:
                out[col] = s

    return out

# ==============================================================================
#  ESTADO INICIAL (como sua versão original)
# ==============================================================================

if 'start_line' not in st.session_state:
    st.session_state.start_line = 1
if 'end_line' not in st.session_state:
    st.session_state.end_line = 5
if 'start_col' not in st.session_state:
    st.session_state.start_col = 1
if 'end_col' not in st.session_state:
    st.session_state.end_col = 10
if 'classification_run' not in st.session_state:
    st.session_state.classification_run = False

try:
    NA_VAL = pd.NA
except:
    NA_VAL = np.nan

# ==============================================================================
#  FUNÇÕES AUXILIARES
# ==============================================================================

def sanitize_filename(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r"[\\/*?\":<>|]", "_", name)
    name = re.sub(r"\s+", "_", name)
    if len(name) > 80: name = name[:80]
    return name or "sheet"


@st.cache_data(show_spinner="Lendo arquivos XLSX…")
def load_multiple_excels(files: List) -> Dict[str, Tuple[str, pd.DataFrame]]:
    """
    Lê múltiplos XLSX, gera estrutura:
        { "Arquivo → Aba": ("Arquivo__Aba", df) }
    """
    merged = {}

    for file in files:
        base = sanitize_filename(file.name.rsplit(".", 1)[0])

        try:
            dfs = pd.read_excel(file, sheet_name=None, header=None, engine="openpyxl")
        except Exception as e:
            st.error(f"Erro ao ler '{file.name}': {e}")
            continue

        for aba, df in dfs.items():
            ui_key = f"{base} → {aba}"
            export_key = f"{base}__{aba}"
            merged[ui_key] = (export_key, df)

    return merged


def normalize_signature(df: pd.DataFrame, r1, r2, c1, c2):
    if df.empty:
        return tuple()

    try:
        sample = df.iloc[r1:r2, c1-1:c2]
    except:
        sample = df.iloc[r1:, c1-1:df.shape[1]]

    signature = []
    for _, row in sample.iterrows():
        clean = []
        for v in row.values:
            s = str(v)
            s = re.sub(r"[^\w]", "", s).lower()
            clean.append(s)
        signature.append("".join(clean))

    return tuple(signature)


@st.cache_resource(show_spinner="Agrupando abas…")
def classify_sheets(tabs_all, r1, r2, c1, c2):
    groups = {}

    for ui_name, (_, df) in tabs_all.items():
        sig = normalize_signature(df, r1, r2, c1, c2)
        groups.setdefault(sig, []).append(ui_name)

    result = {}
    idx = 1

    for sig, lst in groups.items():
        if not sig:
            name = f"Grupo {idx:02d} | Abas Vazias"
        else:
            df_ex = tabs_all[lst[0]][1]
            name = f"Grupo {idx:02d} | {df_ex.shape[1]} Colunas | L{r1}-L{r2} C{c1}-C{c2}"
        result[name] = lst
        idx += 1

    return result


def excel_to_csv_zip(tabs_subset, include_index, sep, encoding):
    buffer = io.BytesIO()

    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for ui_name, (export_key, df) in tabs_subset.items():

            if df.empty:
                continue

            df_export = df.iloc[1:].copy()
            header = df.iloc[0].values
            df_export.columns = list(range(df_export.shape[1]))

            # limpeza nas colunas tipo texto
            for i in range(df_export.shape[1]):
                if df_export.dtypes[i] == "object":
                    df_export[i] = df_export[i].apply(
                        lambda v: NA_VAL if isinstance(v, str) and v.strip()=="" else v
                    )

            df_export.columns = header

            csv_bytes = df_export.to_csv(index=include_index, sep=sep).encode(encoding)

            fname = sanitize_filename(export_key) + ".csv"
            cnt = 1
            while fname in z.namelist():
                fname = sanitize_filename(export_key) + f"_{cnt}.csv"
                cnt += 1

            z.writestr(fname, csv_bytes)

    buffer.seek(0)
    return buffer.read()

# ==============================================================================
#  UI — PRINCIPAL
# ==============================================================================

st.title("📦 XLSX → CSV (ZIP) | Múltiplos Arquivos + Agrupamento Estrutural")
st.markdown("---")

# ------------------ SIDEBAR ------------------

st.sidebar.header("⚙ Exportação")
include_index = st.sidebar.checkbox("Incluir índice", value=False)
encoding = st.sidebar.selectbox("Codificação", ["utf-8", "utf-8-sig", "latin-1"])

sep_label = st.sidebar.selectbox("Separador", ["Vírgula (,)", "Ponto e Vírgula (;)", "Tab (\\t)", "Pipe (|)"])
SEP_MAP = {
    "Vírgula (,)": ",",
    "Ponto e Vírgula (;)": ";",
    "Tab (\\t)": "\t",
    "Pipe (|)": "|",
}
sep = SEP_MAP[sep_label]

st.sidebar.markdown("---")
st.sidebar.header("🔬 Agrupamento")


# ------------------ UPLOAD ------------------

files = st.file_uploader("📚 Envie 1 ou mais arquivos XLSX", type=["xlsx"], accept_multiple_files=True)

if not files:
    st.info("Envie arquivos XLSX para continuar.")
    st.stop()

tabs_all = load_multiple_excels(files)

if not tabs_all:
    st.error("Nenhuma aba encontrada.")
    st.stop()

st.success(f"{len(files)} arquivos → {len(tabs_all)} abas carregadas.")

max_cols = max(df.shape[1] for (_, df) in tabs_all.values())
MAX_LINES = 100

with st.sidebar.form("ranges"):
    st.markdown("**Linhas (0-based)**")
    st.slider("Range", 0, MAX_LINES,
              (st.session_state.start_line, st.session_state.end_line),
              key="slide_l")

    st.markdown("**Colunas (1-based)**")
    st.slider("Range", 1, max_cols,
              (st.session_state.start_col, st.session_state.end_col),
              key="slide_c")

    submit = st.form_submit_button("▶️ Processar")

if submit:
    st.session_state.start_line, st.session_state.end_line = st.session_state.slide_l
    st.session_state.start_col, st.session_state.end_col = st.session_state.slide_c
    st.session_state.classification_run = True

st.markdown("---")

if not st.session_state.classification_run:
    st.info("Ajuste os ranges e clique em **Processar**.")
    st.stop()

# ==============================================================================
# AGRUPAMENTO
# ==============================================================================

r1, r2 = st.session_state.start_line, st.session_state.end_line
c1, c2 = st.session_state.start_col, st.session_state.end_col

groups = classify_sheets(tabs_all, r1, r2, c1, c2)

st.subheader(f"📊 Grupos ({len(groups)}) identificados")
st.caption(f"Baseado em L{r1}-L{r2}, C{c1}-C{c2}")

for group_name, lst in groups.items():
    with st.expander(f"{group_name} — {len(lst)} abas"):

        # lista de abas
        st.markdown("### Abas:")
        st.markdown("\n".join([f"- **{ui}**" for ui in lst]))

        # preview
        example_ui = lst[0]
        export_key, df_ex = tabs_all[example_ui]

        if not df_ex.empty:
            st.markdown("### Cabeçalho (linha 0)")
            head_df = pd.DataFrame({"Colunas": df_ex.iloc[0].fillna("").tolist()})
            st.dataframe(make_safe_display_df(head_df))

            st.markdown("### Prévia (range selecionado)")
            preview = df_ex.iloc[r1:r2, c1-1:c2]
            st.dataframe(make_safe_display_df(preview))

        subset = {name: tabs_all[name] for name in lst}

        if st.button(f"Gerar ZIP — {group_name}", key=f"btn_{group_name}"):
            zip_bytes = excel_to_csv_zip(subset, include_index, sep, encoding)
            fname = sanitize_filename(group_name) + ".zip"

            st.download_button("⬇️ Download ZIP",
                               data=zip_bytes,
                               file_name=fname,
                               mime="application/zip",
                               key=f"dl_{group_name}")

st.markdown("---")
