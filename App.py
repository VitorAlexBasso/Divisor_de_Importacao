import streamlit as st
import pandas as pd
import os
from io import BytesIO
import zipfile
from datetime import datetime

# --- Sanidade do ambiente: precisamos de openpyxl para ler .xlsx/.xlsm
try:
    import openpyxl  # noqa: F401
except Exception:
    st.error(
        "Dependência 'openpyxl' não está instalada no servidor. "
        "Confira seu requirements.txt (deve conter 'openpyxl')."
    )
    st.stop()

# ----------------------------
# Config da aplicação
# ----------------------------
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

# ----------------------------
# Leitura com cache
# ----------------------------
@st.cache_data(show_spinner=False)
def _read_csv_cached(file_bytes, **kwargs):
    from io import BytesIO
    return pd.read_csv(BytesIO(file_bytes), **kwargs)

@st.cache_data(show_spinner=False)
def _read_excel_cached(file_bytes, ext):
    from io import BytesIO
    if ext in (".xlsx", ".xlsm"):
        return pd.read_excel(BytesIO(file_bytes), engine="openpyxl")
    raise ValueError("Formato de arquivo não suportado.")

def process_file(uploaded_file, csv_sep=",", csv_encoding=None):
    try:
        ext = os.path.splitext(uploaded_file.name)[1].lower()
        file_bytes = uploaded_file.getvalue()

        if ext == ".csv":
            return _read_csv_cached(file_bytes, sep=csv_sep, encoding=csv_encoding)
        elif ext in (".xlsx", ".xlsm"):
            return _read_excel_cached(file_bytes, ext)
        else:
            st.error("Formato não suportado. Use .xlsx, .xlsm ou .csv.")
            return None
    except Exception as e:
        st.error(f"Erro na leitura do arquivo: {e}")
        return None

# ----------------------------
# Escrita: divide e zipa sem materializar tudo em memória
# ----------------------------
def stream_zip_from_df(df, base_name, chunk_size, output_format):
    zip_buffer = BytesIO()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    total_rows = len(df)
    num_chunks = (total_rows + chunk_size - 1) // chunk_size

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        progress = st.progress(0, text="Gerando arquivos…")
        for idx, start in enumerate(range(0, total_rows, chunk_size), start=1):
            part = df.iloc[start:start + chunk_size]
            with BytesIO() as out:
                if output_format == "Excel (.xlsx)":
                    # Tenta XlsxWriter (mais rápido); se não houver, cai para openpyxl.
                    try:
                        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
                            part.to_excel(writer, index=False)
                    except Exception:
                        with pd.ExcelWriter(out, engine="openpyxl") as writer:
                            part.to_excel(writer, index=False)
                    filename = f"{base_name}_parte_{idx}_{timestamp}.xlsx"
                else:
                    part.to_csv(out, index=False)
                    filename = f"{base_name}_parte_{idx}_{timestamp}.csv"
                zf.writestr(filename, out.getvalue())
            progress.progress(idx / num_chunks, text=f"Gerando arquivos… {idx}/{num_chunks}")

    zip_buffer.seek(0)
    return zip_buffer, num_chunks, timestamp

def stream_zip_from_csv_file(uploaded_file, base_name, chunk_size, output_format, sep=",", encoding=None):
    # Lê o CSV em chunks e grava direto no ZIP (não carrega tudo)
    zip_buffer = BytesIO()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    reader = pd.read_csv(uploaded_file, chunksize=chunk_size, sep=sep, encoding=encoding)

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        count = 0
        for count, chunk in enumerate(reader, start=1):
            with BytesIO() as out:
                if output_format == "Excel (.xlsx)":
                    try:
                        with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
                            chunk.to_excel(writer, index=False)
                    except Exception:
                        with pd.ExcelWriter(out, engine="openpyxl") as writer:
                            chunk.to_excel(writer, index=False)
                    filename = f"{base_name}_parte_{count}_{timestamp}.xlsx"
                else:
                    chunk.to_csv(out, index=False)
                    filename = f"{base_name}_parte_{count}_{timestamp}.csv"
                zf.writestr(filename, out.getvalue())
    zip_buffer.seek(0)
    return zip_buffer, count, timestamp

# ----------------------------
# UI principal
# ----------------------------
st.title("📊 Divisor de Planilhas PRO")
st.markdown("Divida planilhas grandes em partes menores de forma eficiente. Mantém cabeçalho e estrutura.")

with st.expander("⚙️ Configurações Avançadas"):
    chunk_size = st.number_input(
        "Linhas por arquivo:",
        min_value=100, max_value=10000, value=5000, step=100,
        help="Defina quantas linhas cada arquivo dividido deve ter"
    )
    output_format = st.radio("Formato de saída:", ("Excel (.xlsx)", "CSV (.csv)"), index=0)
    csv_sep = st.text_input("Separador (CSV)", value=",")
    csv_encoding = st.text_input("Encoding (CSV)", value="")

uploaded_file = st.file_uploader(
    "Carregue sua planilha (Excel ou CSV)",
    type=["xlsx", "xlsm", "csv"],  # .xls removido
    help="Arquivos grandes serão automaticamente divididos"
)

if uploaded_file:
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    base_name = os.path.splitext(uploaded_file.name)[0]

    if ext == ".csv":
        use_encoding = csv_encoding if csv_encoding.strip() else None
        st.info("Arquivo CSV detectado. Usando processamento em streaming para melhor desempenho.")
        if st.button("🔀 Dividir Planilha", type="primary"):
            with st.spinner(f"Dividindo CSV em partes de {chunk_size} linhas..."):
                zip_buffer, num_parts, timestamp = stream_zip_from_csv_file(
                    uploaded_file=uploaded_file,
                    base_name=base_name,
                    chunk_size=chunk_size,
                    output_format=output_format,
                    sep=csv_sep,
                    encoding=use_encoding
                )
            st.success(f"✅ Planilha dividida em {num_parts} partes!")
            st.download_button(
                label="⬇️ Baixar Partes (ZIP)",
                data=zip_buffer,
                file_name=f"{base_name}_dividido_{timestamp}.zip",
                mime="application/zip"
            )
    else:
        with st.spinner("Analisando arquivo..."):
            df = process_file(uploaded_file, csv_sep, csv_encoding if csv_encoding.strip() else None)

        if df is not None:
            st.success(f"✅ Arquivo carregado com sucesso! Total de linhas: {len(df):,}")
            if len(df) <= chunk_size:
                st.warning(f"⚠️ A planilha tem menos de {chunk_size} linhas e não será dividida.")
            else:
                num_chunks = (len(df) + chunk_size - 1) // chunk_size
                st.info(f"🔢 Serão gerados {num_chunks} arquivos com ~{chunk_size} linhas cada")
                if st.button("🔀 Dividir Planilha", type="primary"):
                    with st.spinner(f"Dividindo planilha em partes de {chunk_size} linhas..."):
                        zip_buffer, num_parts, timestamp = stream_zip_from_df(
                            df=df,
                            base_name=base_name,
                            chunk_size=chunk_size,
                            output_format=output_format
                        )
                    st.success(f"✅ Planilha dividida em {num_parts} partes!")
                    st.download_button(
                        label="⬇️ Baixar Partes (ZIP)",
                        data=zip_buffer,
                        file_name=f"{base_name}_dividido_{timestamp}.zip",
                        mime="application/zip",
                        help="Clique para baixar o arquivo ZIP com todas as partes"
                    )
                    st.subheader("📊 Estatísticas da Divisão")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Total de Linhas", f"{len(df):,}")
                    col2.metric("Partes Criadas", num_parts)
                    col3.metric("Linhas por Parte", f"{chunk_size:,}")
