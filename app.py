import streamlit as st
from pathlib import Path
from preencher_planilha import (
    extract_lines_from_pdf,
    parse_items,
    build_rows,
    generate_excel_bytes
)

# Configuração da Página
st.set_page_config(page_title="Processador de Planos", layout="wide")

st.title("📄 Extrator de Itens - Plano de Aplicação")
st.info("Este script remove duplicatas (ex: Item 30) e preenche a planilha modelo automaticamente.")

# O modelo de planilha deve estar na mesma pasta
TEMPLATE_FILE = "Planilha_Base.xlsx" 

uploaded_pdf = st.file_uploader("Selecione o PDF do Plano", type=["pdf"])

if uploaded_pdf and st.button("Processar e Gerar Excel"):
    try:
        with st.spinner("Lendo PDF e extraindo dados..."):
            # 1. Extração
            lines = extract_lines_from_pdf(uploaded_pdf)
            
            # 2. Parsing (Com trava de duplicatas)
            items = parse_items(lines)
            
            if not items:
                st.warning("Nenhum item encontrado no PDF. Verifique o formato.")
            else:
                # 3. Formatação
                rows = build_rows(items)
                
                # 4. Excel
                if Path(TEMPLATE_FILE).exists():
                    excel_data = generate_excel_bytes(TEMPLATE_FILE, rows)
                    
                    st.success(f"Processado com sucesso! {len(items)} itens encontrados.")
                    
                    st.download_button(
                        label="📥 Baixar Planilha Preenchida",
                        data=excel_data,
                        file_name="Planilha_Preenchida_Final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error(f"Erro: Arquivo '{TEMPLATE_FILE}' não encontrado na pasta.")
                    
    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
