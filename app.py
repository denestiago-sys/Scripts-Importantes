import base64
import streamlit as st
from io import BytesIO
from pathlib import Path
from openpyxl.utils import get_column_letter
import openpyxl

# Importação sincronizada com as funções reais do preencher_planilha.py
from preencher_planilha import (
    extract_lines_from_pdf_file,
    parse_items,
    build_rows,
    generate_excel_bytes,
    get_template_header_info
)

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "Planilha Base.xlsx" 

st.set_page_config(page_title="Gerador FAF", page_icon="📄", layout="centered")

st.title("Gerador de Planilha de Itens - FAF")
st.markdown("Extração de dados do Plano de Aplicação (CE) para Excel.")

uploaded_file = st.file_uploader("Selecione o PDF", type=["pdf"])

if st.button("Processar e Gerar Planilha", type="primary", disabled=uploaded_file is None):
    if not TEMPLATE_PATH.exists():
        st.error(f"Arquivo '{TEMPLATE_PATH.name}' não encontrado no diretório.")
    else:
        try:
            with st.status("Trabalhando no PDF...", expanded=True) as status:
                status.write("Lendo linhas do documento...")
                lines = extract_lines_from_pdf_file(uploaded_file)
                
                status.write("Identificando itens e corrigindo referências...")
                parsed_items = parse_items(lines)
                
                if not parsed_items:
                    st.error("Nenhum item foi encontrado. Verifique se o PDF é o padrão oficial.")
                else:
                    status.write("Mapeando colunas e valores financeiros...")
                    _, header_map = get_template_header_info(TEMPLATE_PATH)
                    rows = build_rows(parsed_items, header_map)
                    
                    status.write("Finalizando Excel...")
                    excel_bytes = generate_excel_bytes(TEMPLATE_PATH, rows, header_map)
                    
                    st.success(f"Pronto! {len(parsed_items)} itens extraídos com sucesso.")
                    st.download_button(
                        label="📥 Baixar Planilha Preenchida",
                        data=excel_bytes,
                        file_name="Planilha_Preenchida.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"Erro ao processar: {e}")
