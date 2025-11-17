# 📄 XLSX → CSV Processor (Streamlit)

Este aplicativo permite converter facilmente um arquivo **Excel
(.xlsx)** em vários arquivos **CSV**, um para cada aba do arquivo. Após
a conversão, o app gera um **arquivo ZIP** contendo todos os CSVs,
pronto para download.

Ideal para quem precisa extrair dados rapidamente de planilhas
complexas, automatizar conversões ou preparar material para análises em
ferramentas que aceitam apenas CSV.

------------------------------------------------------------------------

## 🚀 Funcionalidades

-   **Upload direto do navegador** de arquivos `.xlsx`\
-   **Leitura automática de todas as abas** do Excel\
-   **Conversão de cada aba para um arquivo CSV individual**
-   **Sanitização automática dos nomes de arquivo**
-   **Seleção opcional das abas que deseja exportar**
-   **Pré-visualização das primeiras linhas de cada aba**
-   **Geração de um ZIP único** contendo todos os CSVs
-   **Opções configuráveis**:
    -   Incluir ou não o índice nos CSVs\
    -   Separador do CSV (`,`, `;`, `tab`)
    -   Codificação (`utf-8`, `utf-8-sig`, `latin-1`)

------------------------------------------------------------------------

## 📥 Como usar

1.  Inicie o app:

    ``` bash
    streamlit run app_xlsx_to_zip.py
    ```

2.  Acesse o navegador (normalmente http://localhost:8501).

3.  Faça o **upload do arquivo .xlsx**.

4.  Opcionalmente, selecione:

    -   Quais abas deseja exportar\
    -   Separador e codificação\
    -   Se deseja incluir o índice

5.  Clique em **Gerar ZIP com CSVs** e baixe o arquivo resultante.

------------------------------------------------------------------------

## 🧩 Dependências

Instale as dependências necessárias:

``` bash
pip install streamlit pandas openpyxl
```

(O app não precisa de outras bibliotecas externas além dessas.)

------------------------------------------------------------------------

## 📦 Estrutura gerada

Após a conversão, o ZIP conterá arquivos nomeados como:

    <nome_da_aba>.csv

Caso existam conflitos ou nomes inválidos, o app ajusta automaticamente.

------------------------------------------------------------------------

## 📝 Observações

-   Apenas arquivos `.xlsx` são suportados.\
-   Abas com nomes muito longos ou caracteres especiais serão
    sanitizadas.\
-   A pré-visualização mostra até 50 linhas por aba.

------------------------------------------------------------------------

## 📚 Licença

Licença livre para uso e modificação.
