import streamlit as st
from pathlib import Path
from preencher_planilha import (
    extract_lines_from_pdf,
    parse_items,
    build_rows,
    generate_excel_bytes
)

st.set_page_config(page_title="Processador de Planos", layout="wide")

st.title("📄 Extrator de Itens - Plano de Aplicação")

# CONFIGURAÇÃO DO NOME DO ARQUIVO
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_NAME = "Planilha Base.xlsx" # Nome atualizado com espaço
TEMPLATE_PATH = BASE_DIR / TEMPLATE_NAME

uploaded_pdf = st.file_uploader("Selecione o PDF do Plano", type=["pdf"])

if uploaded_pdf and st.button("Processar e Gerar Excel"):
    if not TEMPLATE_PATH.exists():
        st.error(f"Erro: O arquivo '{TEMPLATE_NAME}' não foi encontrado na pasta.")
        st.info(f"Caminho esperado: {TEMPLATE_PATH}")
    else:
        try:
            with st.spinner("Processando (Removendo duplicados e preenchendo)..."):
                lines = extract_lines_from_pdf(uploaded_pdf)
                items = parse_items(lines)
                
                if not items:
                    st.warning("Nenhum item detetado no PDF.")
                else:
                    rows = build_rows(items)
                    excel_data = generate_excel_bytes(TEMPLATE_PATH, rows)
                    
                    st.success(f"Sucesso! {len(items)} itens extraídos sem duplicatas.")
                    st.download_button(
                        label="📥 Baixar Planilha Preenchida",
                        data=excel_data,
                        file_name="Planilha_Finalizada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"Erro ao processar: {e}")
