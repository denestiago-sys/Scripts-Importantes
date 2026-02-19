import streamlit as st
from pathlib import Path
from preencher_planilha import (
    extract_lines_from_pdf_file,
    parse_items,
    build_rows,
    generate_excel_bytes,
    get_template_header_info
)

TEMPLATE_PATH = Path(__file__).parent / "Planilha Base.xlsx"

st.set_page_config(page_title="Gerador FAF", layout="centered")
st.title("Gerador de Planilha FAF - Ceará")

uploaded_file = st.file_uploader("Arraste o PDF aqui", type=["pdf"])

if st.button("Processar Plano", type="primary", disabled=not uploaded_file):
    if not TEMPLATE_PATH.exists():
        st.error("Arquivo 'Planilha Base.xlsx' não encontrado no GitHub.")
    else:
        try:
            with st.spinner("Extraindo dados..."):
                lines = extract_lines_from_pdf_file(uploaded_file)
                items = parse_items(lines)
                _, header_map = get_template_header_info(TEMPLATE_PATH)
                rows = build_rows(items, header_map)
                excel = generate_excel_bytes(TEMPLATE_PATH, rows, header_map)
                
                st.success(f"{len(items)} itens processados!")
                st.download_button("📥 Baixar Planilha", excel, "Planilha_Preenchida.xlsx")
        except Exception as e:
            st.error(f"Erro: {e}")
