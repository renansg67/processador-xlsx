# app_xlsx_to_zip.py (Versão Final com Estabilidade de Slider)
import streamlit as st
import pandas as pd
import io
import zipfile
import re
from typing import Dict, List, Tuple
import numpy as np 

# Define o valor de NA (Not Available) para strings vazias após strip
try:
    NA_VAL = pd.NA
except AttributeError:
    NA_VAL = np.nan 

st.set_page_config(page_title="XLSX → CSV (zip) | Agrupamento de Estruturas", layout="wide")

# Inicializa o estado da sessão (Session State) com valores padrão para o range
# Linhas: Início é 1 (segunda linha do XLSX, primeiro dado) e Fim é 4 (quinta linha do XLSX)
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

# ==============================================================================
# ---------- FUNÇÕES DE BACKEND (CACHeadas) ----------
# ==============================================================================

def sanitize_filename(name: str) -> str:
    """Remove caracteres perigosos e espaços extras, limita tamanho."""
    name = str(name).strip()
    name = re.sub(r"[\\/*?\":<>|]", "_", name)
    name = re.sub(r"\s+", "_", name)
    name = re.sub(r"__+", "_", name)
    if len(name) > 80:
        name = name[:80]
    return name or "sheet"

@st.cache_data(show_spinner="Lendo arquivo XLSX. Isso pode levar um tempo para grande volume...")
def load_excel_data(uploaded_file) -> Dict[str, pd.DataFrame]:
    """Carrega todas as abas SEM CABEÇALHO para permitir análise posterior."""
    return pd.read_excel(uploaded_file, sheet_name=None, header=None, engine="openpyxl")


def normalize_content_signature(df: pd.DataFrame, start_line: int, end_line: int, start_col: int, end_col: int) -> Tuple[str, ...]:
    """
    Cria uma 'assinatura' única baseada nos valores dentro do range de linhas e colunas.
    Linhas: Usa o índice do DataFrame (0-based) diretamente.
    Colunas: Converte de Excel (base 1) para Pandas (base 0).
    """
    # Se o DF estiver vazio, retorna uma tupla vazia para que abas vazias sejam agrupadas juntas
    if df.empty:
        return tuple()

    # Linhas: usa índices do DataFrame diretamente (iloc[start:end] exclui o fim)
    start_row_idx = start_line 
    end_row_idx = end_line 
    
    # Colunas: Conversão para índice Pandas (0-based)
    start_col_idx = start_col - 1
    end_col_idx = end_col
    
    # Limita o DF ao range de linhas e COLUNAS definido.
    try:
        sample_df = df.iloc[start_row_idx:end_row_idx, start_col_idx:end_col_idx]
    except IndexError:
        # Pega o máximo que puder dentro do DF
        sample_df = df.iloc[start_row_idx:, start_col_idx:df.shape[1]] 

    signature_list = []
    
    for _, row in sample_df.iterrows():
        row_signature = []
        for value in row.values:
            s = str(value)
            s = re.sub(r'[^\w]', '', s).lower().strip() 
            row_signature.append(s)
        
        signature_list.append("".join(row_signature))
        
    return tuple(signature_list)

@st.cache_resource(show_spinner="Gerando Assinaturas de Conteúdo e Agrupando...")
def classify_sheets_by_content_range(dfs: Dict[str, pd.DataFrame], start_line: int, end_line: int, start_col: int, end_col: int) -> Dict[str, List[str]]:
    """
    Agrupa abas com base na similaridade do conteúdo dentro do range de linhas e colunas.
    """
    groups: Dict[Tuple[str, ...], List[str]] = {}
    
    for sheet_name, df in dfs.items():
        # Cria a assinatura baseada no range de linhas e colunas
        content_signature = normalize_content_signature(df, start_line, end_line, start_col, end_col) 
        
        if content_signature not in groups:
            groups[content_signature] = []
        groups[content_signature].append(sheet_name)
    
    ui_classes: Dict[str, List[str]] = {}
    
    for signature, sheets in groups.items():
        # Se a assinatura for vazia (aba vazia), use um fallback
        if not signature:
            col_count = 0
            class_name = f"Grupo {len(ui_classes) + 1:02d} | Colunas Totais: 0 | ABAS VAZIAS"
        else:
            col_count = dfs[sheets[0]].shape[1]
            # Cria o nome da Classe Estrutural com o range visível
            class_name = f"Grupo {len(ui_classes) + 1:02d} | Colunas Totais: {col_count} | Índices L{start_line}-L{end_line} C{start_col}-C{end_col}"

        ui_classes[class_name] = sheets
        
    return ui_classes

def excel_to_csv_bytes(dfs: Dict[str, pd.DataFrame], include_index: bool, sep: str, encoding: str) -> bytes:
    """Cria um zip em memória contendo os CSVs de cada DataFrame e retorna bytes, aplicando limpeza."""
    
    def safe_strip_and_replace(value):
        """Função que limpa espaços apenas se o valor for uma string, preservando números."""
        if isinstance(value, str):
            stripped = value.strip()
            return NA_VAL if stripped == '' else stripped
        return value
        
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for sheet_name, df in dfs.items():
            if df.empty:
                continue # Pula abas vazias
                
            safe_name = sanitize_filename(sheet_name)
            
            # 1. Cria a cópia do DataFrame a ser exportado
            df_to_export = df.iloc[1:].copy() # Exporta do índice 1 em diante (dados)
            original_header_names = df.iloc[0].values # Usa o índice 0 (cabeçalho)
            
            # 2. LIMPEZA DE DADOS SEGURA (Usando índice de coluna)
            num_cols = df_to_export.shape[1]
            temp_int_columns = list(range(num_cols))
            df_to_export.columns = temp_int_columns
            
            object_cols_indices = [
                i for i, dtype in enumerate(df_to_export.dtypes) if dtype == 'object'
            ]
            
            for i in object_cols_indices:
                df_to_export.iloc[:, i] = df_to_export.iloc[:, i].apply(safe_strip_and_replace)
            
            df_to_export.columns = original_header_names
            
            # 3. Exporta para CSV
            csv_data = df_to_export.to_csv(index=include_index, sep=sep, encoding=encoding)
            csv_bytes = csv_data.encode(encoding)

            # 4. Garante nomes únicos no zip
            entry_name = f"{safe_name}.csv"
            counter = 1
            while entry_name in zf.namelist():
                entry_name = f"{safe_name}_{counter}.csv"
                counter += 1
            zf.writestr(entry_name, csv_bytes)
            
    buffer.seek(0)
    return buffer.read()

# ==============================================================================
# ---------- UI - PAINEL PRINCIPAL ----------
# ==============================================================================

st.title("🔢 XLSX → CSVs → ZIP | Agrupamento de Abas por Estrutura")
st.markdown("---")

# 1. Configurações de Exportação (Sidebar)
st.sidebar.header("🛠️ Parâmetros de Exportação")
include_index = st.sidebar.checkbox("Incluir Índice (Index)", value=False, help="Adiciona a numeração de linha do DataFrame como primeira coluna no CSV.")
encoding = st.sidebar.selectbox("Codificação", options=["utf-8", "utf-8-sig", "latin-1"], index=0, help="Codificação de caracteres do arquivo CSV.")

# Mapeamento de separadores para corrigir a exibição do \t
SEPARATOR_MAP = {
    "Vírgula (,) - Padrão": ",",
    "Ponto e Vírgula (;)": ";",
    "Tabulação (\\t)": "\t", 
    "Pipe (|)": "|",
}

sep_options_display = list(SEPARATOR_MAP.keys())
# Define a vírgula como padrão
default_index = sep_options_display.index("Vírgula (,) - Padrão")

selected_sep_display = st.sidebar.selectbox(
    "Separador CSV", 
    options=sep_options_display, 
    index=default_index, 
    help="Caractere delimitador para as colunas do CSV."
)

# Mapeia o nome de exibição para o valor real do separador
sep_to_use = SEPARATOR_MAP[selected_sep_display]

st.sidebar.markdown("---")
st.sidebar.header("🔬 Parâmetros de Análise")

# 2. Upload
uploaded_file = st.file_uploader("📂 Upload do Arquivo .xlsx", type=["xlsx"], accept_multiple_files=False)

if uploaded_file is None:
    st.info("Aguardando upload de um arquivo Excel (.xlsx).")
    st.stop()

# 3. Leitura CCACHEADA e Cálculo dos Limites
try:
    dfs_all = load_excel_data(uploaded_file)
    
    # Filtra abas completamente vazias (DataFrame com 0 linhas e 0 colunas)
    dfs_all = {name: df for name, df in dfs_all.items() if df.shape[0] > 0 or df.shape[1] > 0}
    num_sheets_total = len(dfs_all)
    
    if num_sheets_total == 0:
        st.warning("O arquivo Excel foi lido, mas não contém abas com conteúdo.")
        st.stop()
        
    st.success(f"Arquivo **{uploaded_file.name}** lido com sucesso ({num_sheets_total} abas com conteúdo).")
    
    # Determina o máximo de colunas para configurar o slider de colunas de forma responsiva
    max_cols = max(df.shape[1] for df in dfs_all.values())
    
    # Fixa o máximo de linhas para o slider em 100
    MAX_SLIDER_LINES = 100 
    
    # Ajusta limites do estado da sessão, se necessário (usando o limite fixo)
    if st.session_state.end_col > max_cols:
        st.session_state.end_col = max_cols
    # Se o valor final do slider estiver acima de 100, ele será corrigido pelo slider abaixo
    if st.session_state.end_line > MAX_SLIDER_LINES:
         st.session_state.end_line = MAX_SLIDER_LINES
             
except Exception as e:
    st.error(f"⚠️ Erro ao processar o arquivo Excel: {e}")
    st.stop()


# 4. Formulário para isolar a lógica de classificação
with st.sidebar.form(key='classification_form'):
    
    # RANGE DE LINHAS (Baseado em ÍNDICE 0)
    st.markdown(f"**Intervalo de Linhas para Comparação (L)**")
    
    # 🚨 ALTERAÇÃO: Removemos a atribuição de variável local para o retorno do slider.
    st.slider(
        "Índice de Linhas (Início / Fim)", 
        min_value=0, 
        max_value=MAX_SLIDER_LINES, 
        value=(st.session_state.start_line, st.session_state.end_line), 
        key='line_slider', # O valor submetido será acessado via st.session_state.line_slider
        help=f"Define o intervalo de linhas usando o **Índice do DataFrame** (0-based). Índice 0 é a linha do Cabeçalho. Máximo Índice para análise: {MAX_SLIDER_LINES}."
    )
    # Lemos os valores submetidos (ou atuais) do Session State associado à chave
    slider_start_line, slider_end_line = st.session_state.line_slider 

    # RANGE DE COLUNAS (Limite Máximo é Dinâmico)
    st.markdown(f"**Intervalo de Colunas para Comparação (C)**")
    
    # 🚨 ALTERAÇÃO: Removemos a atribuição de variável local para o retorno do slider.
    st.slider(
        "Colunas (Excel: Início / Fim)", 
        min_value=1, 
        max_value=max_cols, 
        value=(st.session_state.start_col, st.session_state.end_col), 
        key='col_slider', # O valor submetido será acessado via st.session_state.col_slider
        help=f"Define o intervalo de colunas (Excel: 1 a {max_cols})."
    )
    # Lemos os valores submetidos (ou atuais) do Session State associado à chave
    slider_start_col, slider_end_col = st.session_state.col_slider 

    # Validação
    range_valid = True
    if slider_start_line >= slider_end_line or slider_start_col >= slider_end_col:
        st.error("O início de qualquer intervalo deve ser menor que o fim.")
        range_valid = False
    
    submit_button = st.form_submit_button(label='▶️ Processar Agrupamento')


# 5. Lógica de Agrupamento (Execução Condicional)
if submit_button:
    if range_valid:
        # Atualiza o Session State com os valores do slider
        # Usamos os valores lidos das chaves (que contêm os valores submetidos)
        st.session_state.start_line = slider_start_line
        st.session_state.end_line = slider_end_line
        st.session_state.start_col = slider_start_col
        st.session_state.end_col = slider_end_col
        st.session_state.classification_run = True
    else:
        st.error("Corrija os parâmetros de intervalo e tente novamente.")
        st.session_state.classification_run = False

st.markdown("---")

if st.session_state.classification_run:
    current_start_line = st.session_state.start_line
    current_end_line = st.session_state.end_line
    current_start_col = st.session_state.start_col
    current_end_col = st.session_state.end_col
    
    st.subheader(f"🧱 Grupos de Estrutura de Dados")
    st.caption(f"Agrupamento baseado nos Índices L{current_start_line} até L{current_end_line}, Colunas C{current_start_col} até C{current_end_col}")

    # Chamada da função de classificação
    classified_sheets = classify_sheets_by_content_range(
        dfs_all, 
        start_line=current_start_line, 
        end_line=current_end_line,
        start_col=current_start_col,
        end_col=current_end_col
    )
    class_keys = list(classified_sheets.keys())

    st.info(f"O arquivo foi agrupado em **{len(class_keys)}** Grupos de Estrutura de Dados.")

    # Itera sobre CADA classe e permite a exportação INDIVIDUAL
    for class_name in class_keys:
        sheet_names_in_class = classified_sheets[class_name]
        df_example = dfs_all[sheet_names_in_class[0]]
        
        # FIX para IndexError: Verifica se o DF de exemplo não está vazio
        if df_example.empty:
            actual_header = ["(Aba vazia - Não há cabeçalho ou dados)"]
            preview_df = pd.DataFrame()
            st.warning(f"Aba de Exemplo ('{sheet_names_in_class[0]}') está vazia. Não há pré-visualização ou cabeçalho para esta classe.")
        else:
            actual_header = df_example.iloc[0].fillna('').tolist()
            
            # Lógica para pré-visualização
            preview_start_row_idx = current_start_line 
            preview_end_row_idx = current_end_line
            preview_start_col_idx = current_start_col - 1
            preview_end_col_idx = current_end_col
            
            # Tenta pegar o preview com iloc
            preview_df = df_example.iloc[
                preview_start_row_idx:preview_end_row_idx, 
                preview_start_col_idx:preview_end_col_idx
            ]
            
        with st.expander(f"**{class_name}** | {len(sheet_names_in_class)} Abas"):
            st.markdown(f"**Abas neste grupo:** {', '.join(sheet_names_in_class)}")
            
            st.markdown(f"**Cabeçalho (Índice 0 do DF):**")
            st.dataframe(pd.DataFrame({'Coluna': actual_header}), width="stretch")
            
            # Exibe o preview apenas se houver dados para exibir
            if not preview_df.empty:
                st.markdown(f"**Pré-visualização do Range de Comparação:**")
                st.dataframe(preview_df, width="stretch") 

            # Botão de exportação específico para o GRUPO
            class_export_sheets = {name: dfs_all[name] for name in sheet_names_in_class}
            
            if st.button(f"Gerar ZIP para {class_name}", key=f"btn_export_{class_name}"):
                try:
                    with st.spinner(f"Processando o Grupo {class_name} e limpando dados..."):
                        # O excel_to_csv_bytes já tem a lógica de pular DFs vazios, mas garantimos aqui
                        zip_bytes = excel_to_csv_bytes(class_export_sheets, include_index=include_index, sep=sep_to_use, encoding=encoding)
                        
                    zip_name = f"EXPORT_{class_name.replace('| ', '_').replace(':', '')}_{len(sheet_names_in_class)}_Abas.zip"
                    
                    st.download_button(
                        label=f"⬇️ Baixar ZIP do Grupo {class_name}",
                        data=zip_bytes,
                        file_name=zip_name,
                        mime="application/zip",
                        key=f"dl_{class_name}"
                    )
                    st.success(f"ZIP do Grupo {class_name} pronto para download! (Separador: **{sep_to_use}**)")
                except Exception as e:
                    st.error(f"❌ Falha Crítica ao criar o ZIP do Grupo {class_name}: {e}")
else:
    st.info("Ajuste os parâmetros de **Intervalo de Linhas (Índice)** e **Colunas** na barra lateral e clique em **Processar Agrupamento** para iniciar a análise e visualizar os grupos.")

st.markdown("---")