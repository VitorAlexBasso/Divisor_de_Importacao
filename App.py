import streamlit as st
import pandas as pd
import os
from io import BytesIO
import zipfile
from datetime import datetime

# Configurações da aplicação
def config_app():
    st.set_page_config(
        page_title="Divisor de Planilhas PRO",
        page_icon="📊",
        layout="centered",
        menu_items={
            'Get Help': 'https://github.com/seu-usuario/divisor-planilhas-streamlit',
            'Report a bug': "https://github.com/seu-usuario/divisor-planilhas-streamlit/issues",
            'About': "### Divisor de Planilhas PRO\n\nAplicativo para dividir grandes planilhas em partes menores com eficiência"
        }
    )

# Processamento do arquivo (mantém a lógica do seu app, apenas sem .xls)
def process_file(uploaded_file):
    try:
        file_extension = os.path.splitext(uploaded_file.name)[1].lower()

        if file_extension == '.csv':
            # OBS: do jeito original, lê tudo de uma vez; para arquivos enormes pode pesar.
            df = pd.read_csv(uploaded_file)

        elif file_extension in ('.xlsx', '.xlsm'):
            # openpyxl lê .xlsx/.xlsm normalmente
            df = pd.read_excel(uploaded_file, engine='openpyxl')

        elif file_extension == '.xls':
            # Dica franca aplicada: sem suporte a .xls.
            st.error("Arquivos .xls não são suportados. Regrave como .xlsx (Excel 2007+) e reenviar.")
            return None

        else:
            st.error("Formato de arquivo não suportado")
            return None

        return df

    except Exception as e:
        st.error(f"Erro na leitura do arquivo: {str(e)}")
        return None

# Divisão da planilha
def split_dataframe(df, chunk_size=5000):
    return [df[i:i + chunk_size] for i in range(0, len(df), chunk_size)]

# Interface principal
def main():
    config_app()

    st.title("📊 Divisor de Planilhas PRO")
    st.markdown("""
    Divida planilhas grandes em partes menores de forma eficiente.
    Mantém todos os formatos e cabeçalhos originais.
    """)

    with st.expander("⚙️ Configurações Avançadas"):
        chunk_size = st.number_input(
            "Linhas por arquivo:",
            min_value=100,
            max_value=10000,
            value=5000,
            step=100,
            help="Defina quantas linhas cada arquivo dividido deve ter"
        )
        output_format = st.radio(
            "Formato de saída:",
            ("Excel (.xlsx)", "CSV (.csv)"),
            index=0
        )

    uploaded_file = st.file_uploader(
        "Carregue sua planilha (Excel ou CSV)",
        type=["xlsx", "xlsm", "csv"],  # <- .xls removido
        help="Arquivos grandes serão automaticamente divididos"
    )

    if uploaded_file:
        with st.spinner("Analisando arquivo..."):
            df = process_file(uploaded_file)

            if df is not None:
                st.success(f"✅ Arquivo carregado com sucesso! Total de linhas: {len(df):,}")

                if len(df) <= chunk_size:
                    st.warning(f"⚠️ A planilha tem menos de {chunk_size} linhas e não será dividida.")
                else:
                    num_chunks = (len(df) // chunk_size) + (1 if len(df) % chunk_size else 0)
                    st.info(f"🔢 Serão gerados {num_chunks} arquivos com ~{chunk_size} linhas cada")

                    if st.button("🔀 Dividir Planilha", type="primary"):
                        with st.spinner(f"Dividindo planilha em partes de {chunk_size} linhas..."):
                            chunks = split_dataframe(df, chunk_size)
                            zip_buffer = BytesIO()
                            base_name = os.path.splitext(uploaded_file.name)[0]
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

                            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                                for i, chunk in enumerate(chunks, start=1):
                                    with BytesIO() as output:
                                        if output_format == "Excel (.xlsx)":
                                            # Se quiser acelerar, instale XlsxWriter e troque o engine aqui.
                                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                                chunk.to_excel(writer, index=False)
                                            file_name = f"{base_name}_parte_{i}_{timestamp}.xlsx"
                                        else:
                                            chunk.to_csv(output, index=False)
                                            file_name = f"{base_name}_parte_{i}_{timestamp}.csv"

                                        zip_file.writestr(file_name, output.getvalue())

                            zip_buffer.seek(0)
                            st.success(f"✅ Planilha dividida em {len(chunks)} partes!")

                            st.download_button(
                                label="⬇️ Baixar Partes (ZIP)",
                                data=zip_buffer,
                                file_name=f"{base_name}_dividido_{timestamp}.zip",
                                mime="application/zip",
                                help="Clique para baixar o arquivo ZIP com todas as partes"
                            )

                            # Mostrar estatísticas
                            st.subheader("📊 Estatísticas da Divisão")
                            col1, col2, col3 = st.columns(3)
                            col1.metric("Total de Linhas", f"{len(df):,}")
                            col2.metric("Partes Criadas", len(chunks))
                            col3.metric("Linhas por Parte", f"{chunk_size:,}")

if __name__ == "__main__":
    main()
