import streamlit as st
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
import io

# --- FUNÇÕES DE APOIO ---
def extract_lines_from_pdf_file(pdf_file):
    lines = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines.extend(text.split('\n'))
    return lines

def main():
    st.set_page_config(page_title="Conversor Portaria 685", page_icon="📄")
    
    st.title("📄 Conversor Portaria 685")
    st.markdown("Suba o PDF e a planilha modelo para gerar o arquivo preenchido.")

    # Upload dos arquivos
    pdf_file = st.file_uploader("PDF de Entrada", type="pdf")
    excel_file = st.file_uploader("Planilha Modelo", type="xlsx")

    if pdf_file and excel_file:
        if st.button("Processar e Gerar Planilha"):
            try:
                # 1. Extração
                lines = extract_lines_from_pdf_file(pdf_file)
                
                # 2. Processamento (Sua lógica de dados aqui)
                # Exemplo simples: criando um DataFrame com as linhas
                df_pdf = pd.DataFrame(lines, columns=["Conteúdo PDF"])

                # 3. Preparar Excel em memória
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_pdf.to_excel(writer, index=False, sheet_name='Resultado')
                
                st.success("✅ Processamento concluído!")
                
                st.download_button(
                    label="📥 Baixar Resultado",
                    data=output.getvalue(),
                    file_name="portaria_processada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    main()
