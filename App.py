import streamlit as st
import pandas as pd
import os
from io import BytesIO
import zipfile

# Configuração da página
st.set_page_config(
    page_title="Divisor de Planilhas",
    page_icon="📊",
    layout="centered"
)

# Função principal
def main():
    st.title("📊 Divisor de Planilhas")
    st.markdown("""
    Divida sua planilha em arquivos menores de 5.000 linhas cada.
    Mantém todos os formatos e cabeçalhos originais.
    """)

    # Upload do arquivo
    uploaded_file = st.file_uploader(
        "Carregue sua planilha (Excel ou CSV)",
        type=["xlsx", "xls", "csv"],
        help="Arquivos muito grandes serão divididos em partes de 5.000 linhas"
    )

    if uploaded_file is not None:
        # Processar o arquivo
        try:
            # Ler o arquivo
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            # Verificar tamanho
            if len(df) <= 5000:
                st.warning("A planilha tem menos de 5.000 linhas e não será dividida.")
                st.download_button(
                    label="Baixar Planilha Original",
                    data=uploaded_file.getvalue(),
                    file_name=uploaded_file.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                return

            # Configurações
            chunk_size = 5000
            start_row = 1  # Começa da linha 2 (0-indexed)

            # Criar um arquivo ZIP para todos os pedaços
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
                # Dividir o DataFrame
                total_chunks = (len(df) - start_row) // chunk_size + 1
                base_name = os.path.splitext(uploaded_file.name)[0]

                progress_bar = st.progress(0)
                status_text = st.empty()

                for i in range(total_chunks):
                    start = start_row + i * chunk_size
                    end = min(start + chunk_size, len(df))
                    
                    chunk = df.iloc[start:end]
                    
                    # Criar arquivo Excel em memória
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        chunk.to_excel(writer, index=False)
                    
                    # Adicionar ao ZIP
                    file_name = f"{base_name}_parte_{i+1}.xlsx"
                    zip_file.writestr(file_name, output.getvalue())
                    
                    # Atualizar progresso
                    progress = (i + 1) / total_chunks
                    progress_bar.progress(progress)
                    status_text.text(f"Processando: parte {i+1} de {total_chunks}...")

                status_text.text("✅ Processamento concluído!")

            # Botão para download do ZIP
            zip_buffer.seek(0)
            st.download_button(
                label="⬇️ Baixar Todas as Partes (ZIP)",
                data=zip_buffer,
                file_name=f"{base_name}_dividido.zip",
                mime="application/zip"
            )

            st.success(f"Planilha dividida em {total_chunks} partes de até 5.000 linhas cada.")

        except Exception as e:
            st.error(f"Ocorreu um erro: {str(e)}")

# Rodar o app
if __name__ == "__main__":
    main()
