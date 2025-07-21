import streamlit as st
import pandas as pd
import os
from io import BytesIO
import zipfile

# Configurações da aplicação
def config_app():
    st.set_page_config(
        page_title="Divisor de Planilhas",
        page_icon="📊",
        layout="centered",
        menu_items={
            'Get Help': 'https://github.com/seu-usuario/divisor-planilhas-streamlit',
            'Report a bug': "https://github.com/seu-usuario/divisor-planilhas-streamlit/issues",
            'About': "### Divisor de Planilhas\n\nAplicativo para dividir grandes planilhas em partes menores"
        }
    )

# Processamento do arquivo
def process_file(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        return df
    except Exception as e:
        st.error(f"Erro na leitura do arquivo: {str(e)}")
        return None

# Divisão da planilha
def split_dataframe(df, chunk_size=5000):
    chunks = []
    for i in range(0, len(df), chunk_size):
        chunks.append(df[i:i + chunk_size])
    return chunks

# Interface principal
def main():
    config_app()
    
    st.title("📊 Divisor de Planilhas")
    st.markdown("""
    Divida planilhas grandes em partes menores de 5.000 linhas cada.
    Mantém todos os formatos e cabeçalhos originais.
    """)

    uploaded_file = st.file_uploader(
        "Carregue sua planilha (Excel ou CSV)",
        type=["xlsx", "xls", "csv"],
        help="Arquivos com mais de 5.000 linhas serão automaticamente divididos"
    )

    if uploaded_file:
        df = process_file(uploaded_file)
        
        if df is not None:
            if len(df) <= 5000:
                st.warning("⚠️ A planilha tem menos de 5.000 linhas e não será dividida.")
            else:
                with st.spinner("Processando divisão da planilha..."):
                    chunks = split_dataframe(df)
                    zip_buffer = BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for i, chunk in enumerate(chunks):
                            with BytesIO() as output:
                                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                    chunk.to_excel(writer, index=False)
                                zip_file.writestr(
                                    f"{os.path.splitext(uploaded_file.name)[0]}_parte_{i+1}.xlsx",
                                    output.getvalue()
                                )
                    
                    zip_buffer.seek(0)
                    st.success(f"✅ Planilha dividida em {len(chunks)} partes!")
                    
                    st.download_button(
                        label="⬇️ Baixar Partes (ZIP)",
                        data=zip_buffer,
                        file_name="planilha_dividida.zip",
                        mime="application/zip"
                    )

if __name__ == "__main__":
    main()
