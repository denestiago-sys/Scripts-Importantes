import base64
import streamlit as st
import pdfplumber
import pandas as pd
import io
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# --- FUNÇÕES DE LÓGICA (ANTIGO PREENCHER_PLANILHA.PY) ---

def extract_lines_from_pdf_file(pdf_file):
    lines = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines.extend(text.split('\n'))
    return lines

def parse_items(lines):
    """
    Sua lógica de extração de itens do PDF. 
    Aqui deve conter as regras para encontrar Número da Meta, Item, etc.
    """
    items = []
    # Exemplo de estrutura que seu script build_rows espera:
    # for line in lines: ... lógica de captura ...
    # items.append({"Número da Meta Específica": "01", "Número do Item": "1", ...})
    
    # IMPORTANTE: Como não tenho seu parse_items original completo, 
    # certifique-se de que a lógica de captura de dados está aqui dentro.
    return items 

def get_template_header_info(template_path):
    wb = load_workbook(template_path, data_only=True)
    ws = wb.active
    header_map = {}
    # Assume que o cabeçalho está na linha 2 (ajuste se necessário)
    for cell in ws[2]: 
        if cell.value:
            header_map[cell.value] = cell.column
    return wb, header_map

def build_rows(parsed_items, header_map):
    rows = []
    for item in parsed_items:
        row_data = {header: item.get(header, "") for header in header_map.keys()}
        rows.append(row_data)
    return rows

def generate_excel_bytes(template_path, rows, header_map):
    wb = load_workbook(template_path)
    ws = wb.active
    start_row = 3
    for i, row_data in enumerate(rows):
        for header, col_idx in header_map.items():
            ws.cell(row=start_row + i, column=col_idx, value=row_data.get(header))
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- INTERFACE STREAMLIT (SEU CÓDIGO) ---

BASE_DIR = Path(__file__).resolve().parent
# Nota: O arquivo 'Planilha Base.xlsx' precisa estar no seu GitHub!
TEMPLATE_PATH = BASE_DIR / "Planilha Base.xlsx"
LOGO_PATH = BASE_DIR / "Logo.png"

st.set_page_config(page_title="Preenche Planilhas", page_icon="📄", layout="centered")

# [O SEU CSS AQUI - MANTIDO IGUAL]
st.markdown("""<style>...</style>""", unsafe_allow_html=True) 

# Lógica da Logo
logo_b64 = ""
if LOGO_PATH.exists():
    logo_bytes = LOGO_PATH.read_bytes()
    logo_b64 = base64.b64encode(logo_bytes).decode("ascii")

logo_html = f'<div class="logo-wrap"><img src="data:image/png;base64,{logo_b64}" /></div>' if logo_b64 else ""

st.markdown(f'<div class="header">{logo_html}<h1 class="header-title">Gerador de Planilha de Itens - FAF</h1></div>', unsafe_allow_html=True)
st.markdown('<div class="brand-bar"><span></span><span></span><span></span><span></span><span></span></div>', unsafe_allow_html=True)
st.markdown('<p class="app-subtitle">Faça upload do PDF e gere a planilha automaticamente.</p>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("PDF do Plano", type=["pdf"])

if "result" not in st.session_state:
    st.session_state.result = None

if st.button("Processar", type="primary", disabled=uploaded_file is None):
    if not TEMPLATE_PATH.exists():
        st.error(f"Arquivo 'Planilha Base.xlsx' não encontrado no repositório. Suba ele para o GitHub!")
    else:
        try:
            with st.status("Processando PDF...", expanded=True) as status:
                status.write("Lendo PDF")
                lines = extract_lines_from_pdf_file(uploaded_file)

                status.write("Extraindo itens")
                parsed_items = parse_items(lines)
                
                if not parsed_items:
                    status.update(label="Nenhum item encontrado.", state="error")
                    st.error("Nenhum item encontrado no PDF. Verifique o formato do arquivo.")
                else:
                    status.write("Montando planilha")
                    _, header_map = get_template_header_info(TEMPLATE_PATH)
                    rows = build_rows(parsed_items, header_map)
                    excel_bytes = generate_excel_bytes(TEMPLATE_PATH, rows, header_map)

                    # Lógica de resumo (conforme seu original)
                    st.session_state.result = {
                        "rows": rows,
                        "excel_bytes": excel_bytes,
                        "meta_counts": {}, # Adicione lógica de contagem aqui
                        "missing_cells": [],
                        "missing_items_count": 0,
                    }
                    status.update(label="Processamento concluído.", state="complete")
        except Exception as exc:
            st.exception(exc)

# Exibição do Resultado
if st.session_state.result:
    res = st.session_state.result
    st.download_button(
        "Baixar Planilha",
        data=res["excel_bytes"],
        file_name="Planilha de Itens.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
