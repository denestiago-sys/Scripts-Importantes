import base64
import streamlit as st
from io import BytesIO
from pathlib import Path
from openpyxl.utils import get_column_letter
import openpyxl

# Importando apenas o que realmente existe no preencher_planilha.py corrigido
from preencher_planilha import (
    extract_lines_from_pdf_file,
    parse_items,
    build_rows,
    generate_excel_bytes,
    get_template_header_info,
    is_analysis_template_file,
    extract_plan_signature,
    resolve_art_by_plan_rule
)

BASE_DIR = Path(__file__).resolve().parent
# Verifique se o nome do arquivo no seu GitHub é "Planilha Base.xlsx"
TEMPLATE_PATH = BASE_DIR / "Planilha Base.xlsx" 
LOGO_PATH = BASE_DIR / "Logo.png"

st.set_page_config(page_title="Gerador de Planilha FAF", page_icon="📄", layout="centered")

# --- Interface Visual ---
st.title("Gerador de Planilha de Itens - FAF")
st.markdown("Faça upload do PDF do Plano de Aplicação e gere a planilha preenchida automaticamente.")

uploaded_file = st.file_uploader("Upload do PDF", type=["pdf"])

if st.button("Processar Plano", type="primary", disabled=uploaded_file is None):
    if not TEMPLATE_PATH.exists():
        st.error(f"Erro: O arquivo '{TEMPLATE_PATH.name}' não foi encontrado no GitHub.")
    else:
        try:
            with st.status("Processando dados...", expanded=True) as status:
                status.write("Lendo PDF e limpando ruídos...")
                lines = extract_lines_from_pdf_file(uploaded_file)
                
                status.write("Extraindo itens (Corrigindo Item 30 e textos longos)...")
                parsed_items = parse_items(lines)
                
                if not parsed_items:
                    st.error("Nenhum item foi identificado. Verifique se o PDF está no padrão oficial.")
                else:
                    status.write("Mapeando colunas da planilha...")
                    # Obtém o cabeçalho da sua planilha base (Linha 2)
                    _, header_map = get_template_header_info(TEMPLATE_PATH)
                    
                    # Organiza os dados para as colunas
                    rows = build_rows(parsed_items, header_map)
                    
                    status.write("Gerando arquivo Excel...")
                    excel_bytes = generate_excel_bytes(TEMPLATE_PATH, rows, header_map)
                    
                    st.success(f"Sucesso! {len(parsed_items)} itens processados corretamente.")
                    
                    st.download_button(
                        label="📥 Baixar Planilha Preenchida",
                        data=excel_bytes,
                        file_name="Planilha_FAF_Preenchida.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"Ocorreu um erro técnico: {e}")
            st.exception(e)
