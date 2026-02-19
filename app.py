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
st.success("Filtro de precisão para Valores Financeiros (Item 127) ativado.")

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_NAME = "Planilha Base.xlsx"
TEMPLATE_PATH = BASE_DIR / TEMPLATE_NAME

uploaded_pdf = st.file_uploader("Selecione o PDF do Plano", type=["pdf"])

if uploaded_pdf and st.button("Processar e Gerar Excel"):
    if not TEMPLATE_PATH.exists():
        st.error(f"Erro: O arquivo '{TEMPLATE_NAME}' não foi encontrado.")
    else:
        try:
            with st.spinner("Limpando ruídos e corrigindo valores..."):
                lines = extract_lines_from_pdf(uploaded_pdf)
                items = parse_items(lines)
                
                if not items:
                    st.warning("Nenhum item detectado no PDF.")
                else:
                    rows = build_rows(items)
                    excel_data = generate_excel_bytes(TEMPLATE_PATH, rows)
                    
                    st.success(f"Concluído! {len(items)} itens processados com valores corrigidos.")
                    st.download_button(
                        label="📥 Baixar Planilha Final Corrigida",
                        data=excel_data,
                        file_name="Planilha_Final_Sem_Erros.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"Erro ao processar: {e}")
