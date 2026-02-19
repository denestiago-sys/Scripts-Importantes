import streamlit as st
from pathlib import Path
from preencher_planilha import (
    extract_lines_from_pdf_file,
    parse_items,
    build_rows,
    generate_excel_bytes,
    get_template_header_info
)

TEMPLATE_NAME = "Planilha Base.xlsx"
TEMPLATE_PATH = Path(__file__).parent / TEMPLATE_NAME

st.set_page_config(page_title="FAF - Processador", layout="wide")
st.title("📄 Extração de Dados FAF (Padrão Completo)")

uploaded_file = st.file_uploader("Upload do PDF do Plano de Aplicação", type=["pdf"])

if st.button("Gerar Planilha Completa", type="primary") and uploaded_file:
    try:
        with st.status("Processando documento...", expanded=True) as status:
            status.write("Lendo PDF e aplicando Regex...")
            lines = extract_lines_from_pdf_file(uploaded_file)
            
            status.write("Agrupando Itens e Metas...")
            items = parse_items(lines)
            
            if not items:
                st.error("Nenhum item detectado. Verifique o formato do PDF.")
            else:
                status.write(f"Mapeando {len(items)} itens para o Excel...")
                header_row, header_map = get_template_header_info(TEMPLATE_PATH)
                rows = build_rows(items, header_map)
                
                excel_data = generate_excel_bytes(TEMPLATE_PATH, rows, header_map)
                
                status.update(label="Processamento concluído!", state="complete")
                st.success(f"Sucesso: {len(items)} itens extraídos.")
                st.download_button(
                    label="📥 Baixar Planilha Preenchida",
                    data=excel_data,
                    file_name="FAF_Preenchido.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"Erro crítico: {str(e)}")
