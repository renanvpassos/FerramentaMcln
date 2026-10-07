import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Tratador de Planilhas", page_icon="📊", layout="wide")

st.title("📊 Ferramenta de Tratamento de Planilhas")
st.write("Faça o upload da sua planilha de origem para gerar o arquivo formatado conforme as regras estabelecidas.")

uploaded_file = st.file_uploader("Escolha a planilha de origem (Excel ou CSV)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
try:
# Leitura do arquivo conforme extensão
if uploaded_file.name.endswith('.csv'):
df_origem = pd.read_csv(uploaded_file)
else:
df_origem = pd.read_excel(uploaded_file)

    st.success("Planilha carregada com sucesso!")
    with st.expander("Ver prévia da planilha original"):
        st.dataframe(df_origem.head())

    # Normalizar nomes das colunas (remover espaços extras e padronizar maiúsculas/minúsculas para busca)
    df_origem.columns = [str(col).strip() for col in df_origem.columns]

    # Criação do novo DataFrame estruturado
    df_novo = pd.DataFrame()

    # 1. COLUNA A: "INVOICE"
    col_invoice = next((c for c in df_origem.columns if c.upper() == 'INVOICE'), None)
    if col_invoice:
        df_novo['INVOICE'] = df_origem[col_invoice]
    else:
        df_novo['INVOICE'] = ""
        st.warning("Coluna 'INVOICE' não encontrada na planilha de origem.")

    # 2. COLUNA B: "PARTNUMBER" (Coluna D ou "SKU")
    col_sku = next((c for c in df_origem.columns if c.upper() == 'SKU'), None)
    if col_sku:
        df_novo['PARTNUMBER'] = df_origem[col_sku]
    elif len(df_origem.columns) >= 4:
        df_novo['PARTNUMBER'] = df_origem.iloc[:, 3] # Índice 3 corresponde à Coluna D
        st.info("Coluna 'SKU' não encontrada pelo nome exato, utilizando a Coluna D.")
    else:
        df_novo['PARTNUMBER'] = ""
        st.warning("Coluna 'PARTNUMBER' (SKU / Coluna D) não encontrada.")

    # 3. COLUNA C: "QUANTIDADE" ("Qt")
    col_qt = next((c for c in df_origem.columns if c.upper() in ['QT', 'QUANTIDADE']), None)
    if col_qt:
        df_novo['QUANTIDADE'] = df_origem[col_qt]
    else:
        df_novo['QUANTIDADE'] = ""
        st.warning("Coluna 'Qt' (Quantidade) não encontrada.")

    # 4. COLUNA D: "UNIDADE" ("X")
    col_x = next((c for c in df_origem.columns if c.upper() == 'X'), None)
    if col_x:
        df_novo['UNIDADE'] = df_origem[col_x]
    else:
        df_novo['UNIDADE'] = ""
        st.warning("Coluna 'X' (Unidade) não encontrada.")

    # 5. COLUNA E: "PRECOTOTAL" ("Vlr. Total")
    col_vlr = next((c for c in df_origem.columns if 'VLR' in c.upper() and 'TOTAL' in c.upper()), None)
    if col_vlr:
        df_novo['PRECOTOTAL'] = df_origem[col_vlr]
    else:
        df_novo['PRECOTOTAL'] = ""
        st.warning("Coluna 'Vlr. Total' não encontrada.")

    # 6 & 7. COLUNAS F e G: "MOEDA" e "INCOTERMS" baseados no nome do arquivo
    filename_lower = uploaded_file.name.lower()
    is_thelios = "thelios" in filename_lower

    df_novo['MOEDA'] = "220" if is_thelios else "790"
    df_novo['INCOTERMS'] = "EXW" if is_thelios else "CPT"

    # 8. COLUNA H: "PESOUNITARIO" ("Peso+SKU", extraindo apenas números do início)
    col_peso = next((c for c in df_origem.columns if 'PESO' in c.upper() and 'SKU' in c.upper()), None)
    if col_peso:
        def extrair_peso(valor):
            if pd.isna(valor):
                return ""
            match = re.match(r'^([0-9]+(?:[,\.][0-9]+)?)', str(valor).strip())
            return match.group(1) if match else ""
        
        df_novo['PESOUNITARIO'] = df_origem[col_peso].apply(extrair_peso)
    else:
        df_novo['PESOUNITARIO'] = ""
        st.warning("Coluna 'Peso+SKU' não encontrada.")

    st.subheader("Prévia da nova planilha tratada:")
    st.dataframe(df_novo.head())

    # Geração do arquivo Excel para download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_novo.to_excel(writer, index=False, sheet_name='Tratado')
    processed_data = output.getvalue()

    st.download_button(
        label="📥 Baixar Planilha Tratada (Excel)",
        data=processed_data,
        file_name="planilha_tratada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

except Exception as e:
    st.error(f"Ocorreu um erro ao processar o arquivo: {e}")


else:
st.info("Aguardando o upload de um arquivo para iniciar...")
