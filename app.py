import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Processador de Planilhas", page_icon="📊", layout="centered"
)

st.title("📊 Processador de Planilhas")
st.write(
    "Faça o upload da sua planilha de origem para gerar o arquivo tratado."
)

uploaded_file = st.file_uploader(
    "Escolha o arquivo Excel ou CSV", type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
    try:
        # Detecta se é CSV ou Excel
        file_name = uploaded_file.name
        if file_name.endswith(".csv"):
            df_origem = pd.read_csv(uploaded_file)
        else:
            df_origem = pd.read_excel(uploaded_file)

        st.success("Planilha carregada com sucesso!")
        st.write("Prévia dos dados originais:", df_origem.head())

        if st.button("Processar Planilha"):
            df_destino = pd.DataFrame()

            # COLUNA A: INVOICE
            if "INVOICE" in df_origem.columns:
                df_destino["INVOICE"] = df_origem["INVOICE"]
            else:
                st.warning("Coluna 'INVOICE' não encontrada na origem.")
                df_destino["INVOICE"] = ""

            # COLUNA B: PARTNUMBER (Coluna D 'SKU')
            if "SKU" in df_origem.columns:
                df_destino["PARTNUMBER"] = df_origem["SKU"]
            else:
                st.warning("Coluna 'SKU' não encontrada na origem.")
                df_destino["PARTNUMBER"] = ""

            # COLUNA C: QUANTIDADE (Coluna 'Qt')
            if "Qt" in df_origem.columns:
                df_destino["QUANTIDADE"] = df_origem["Qt"]
            else:
                st.warning("Coluna 'Qt' não encontrada na origem.")
                df_destino["QUANTIDADE"] = ""

            # COLUNA D: UNIDADE (Coluna 'X')
            if "X" in df_origem.columns:
                df_destino["UNIDADE"] = df_origem["X"]
            else:
                st.warning("Coluna 'X' não encontrada na origem.")
                df_destino["UNIDADE"] = ""

            # COLUNA E: PRECOTOTAL (Coluna ' Vlr. Total ')
            col_vlr = next(
                (c for c in df_origem.columns if "Vlr. Total" in c), None
            )
            if col_vlr:
                df_destino["PRECOTOTAL"] = df_origem[col_vlr]
            else:
                st.warning("Coluna ' Vlr. Total ' não encontrada na origem.")
                df_destino["PRECOTOTAL"] = ""

            # Regras baseadas no nome do arquivo
            is_thelios = (
                "THELIOS" in file_name.upper() or "THELIOS" in file_name
            )

            # COLUNA F: MOEDA
            df_destino["MOEDA"] = "220" if is_thelios else "790"

            # COLUNA G: INCOTERMS
            df_destino["INCOTERMS"] = "EXW" if is_thelios else "CPT"

            # COLUNA H: PESOUNITARIO (Extrair apenas os números do início da coluna ' Peso+SKU ')
            col_peso = next((c for c in df_origem.columns if "Peso+SKU" in c), None)
            if col_peso:

                def extrair_peso(valor):
                    val_str = str(valor)
                    match = re.match(r"^([\d,\.]+)", val_str)
                    return match.group(1) if match else val_str

                df_destino["PESOUNITARIO"] = df_origem[col_peso].apply(
                    extrair_peso
                )
            else:
                st.warning("Coluna ' Peso+SKU ' não encontrada na origem.")
                df_destino["PESOUNITARIO"] = ""

            st.write("Prévia dos dados tratados:", df_destino.head())

            # Botão de Download do arquivo processado
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_destino.to_excel(writer, index=False, sheet_name="Tratado")
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Baixar Planilha Tratada",
                data=processed_data,
                file_name="planilha_tratada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
