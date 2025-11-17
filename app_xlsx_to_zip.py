# app_xlsx_to_zip.py
import streamlit as st
import pandas as pd
import io
import zipfile
import re
from typing import Dict

st.set_page_config(page_title="XLSX → CSV (zip)", layout="wide")

st.title("📥 XLSX → CSVs → ZIP")
st.write("Faça upload de um arquivo Excel (.xlsx). Cada aba será convertida para CSV e empacotada num ZIP para download.")

# ---------- helpers ----------
def sanitize_filename(name: str) -> str:
    """Remove caracteres perigosos e espaços extras, limita tamanho."""
    name = str(name)
    name = name.strip()
    name = re.sub(r"[\\/*?\":<>|]", "_", name)  # caracteres inválidos em arquivos
    name = re.sub(r"\s+", "_", name)
    name = re.sub(r"__+", "_", name)
    if len(name) > 80:
        name = name[:80]
    return name or "sheet"

def excel_to_csv_bytes(dfs: Dict[str, pd.DataFrame], include_index: bool, sep: str, encoding: str) -> bytes:
    """Cria um zip em memória contendo os CSVs de cada DataFrame e retorna bytes."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for sheet_name, df in dfs.items():
            safe_name = sanitize_filename(sheet_name)
            csv_bytes = df.to_csv(index=include_index, sep=sep, encoding=encoding).encode(encoding)
            # garante extensão .csv e nomes únicos no zip
            entry_name = f"{safe_name}.csv"
            counter = 1
            while entry_name in zf.namelist():
                entry_name = f"{safe_name}_{counter}.csv"
                counter += 1
            zf.writestr(entry_name, csv_bytes)
    buffer.seek(0)
    return buffer.read()

# ---------- UI ----------
uploaded_file = st.file_uploader("Carregue um arquivo .xlsx", type=["xlsx"], accept_multiple_files=False)

if uploaded_file is None:
    st.info("Envie um arquivo Excel (.xlsx) para começar.")
    st.stop()

# Opções
st.sidebar.header("Opções de conversão")
include_index = st.sidebar.checkbox("Incluir índice do DataFrame no CSV", value=False)
sep = st.sidebar.selectbox("Separador do CSV", options=[",", ";", "\t"], index=0, help="Escolha vírgula, ponto-e-vírgula ou tab")
encoding = st.sidebar.selectbox("Codificação", options=["utf-8", "utf-8-sig", "latin-1"], index=0)

# Leitura - tenta carregar todas as abas
try:
    # sheet_name=None lê todas as abas e retorna dict sheet_name -> DataFrame
    dfs_all = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")
except Exception as e:
    st.error(f"Erro ao ler o arquivo Excel: {e}")
    st.stop()

st.success(f"Lidas {len(dfs_all)} abas do arquivo.")
st.markdown("**Visualizar / selecionar abas**")

# mostra checkboxes para seleção de abas
columns = st.columns([3,1])
with columns[0]:
    selected = st.multiselect("Escolha as abas para exportar (deixe vazio = todas):", options=list(dfs_all.keys()))
with columns[1]:
    if st.button("Selecionar todas"):
        selected = list(dfs_all.keys())

# Decide quais abas serão efetivamente exportadas
if selected:
    export_sheets = {k: dfs_all[k] for k in selected}
else:
    export_sheets = dfs_all

# Mostra pré-visualização compacta de cada aba selecionada
for name, df in list(export_sheets.items())[:5]:  # evita mostrar preview de muitas abas de uma vez
    with st.expander(f"Aba: {name} — {df.shape[0]} linhas × {df.shape[1]} colunas"):
        st.dataframe(df.head(50), use_container_width=True)

if len(export_sheets) > 5:
    st.info(f"Mostrando preview das primeiras 5 abas. {len(export_sheets)-5} abas ocultas no preview.")

# Botão para gerar ZIP
st.markdown("---")
if st.button("Gerar ZIP com CSVs"):
    try:
        zip_bytes = excel_to_csv_bytes(export_sheets, include_index=include_index, sep=sep, encoding=encoding)
        zip_name = sanitize_filename(uploaded_file.name.rsplit(".", 1)[0]) + "_csvs.zip"
        st.success("ZIP gerado com sucesso ✅")

        st.download_button(
            label="⬇️ Baixar ZIP com os CSVs",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip"
        )
    except Exception as e:
        st.error(f"Falha ao criar o ZIP: {e}")
else:
    st.write("Clique em **Gerar ZIP com CSVs** para iniciar a conversão e obter o botão de download.")

# Pequeno rodapé com dicas
st.caption("Dica: se algum nome de aba for muito longo ou contiver caracteres especiais, ele será sanitizado para criar nomes de arquivos válidos.")


