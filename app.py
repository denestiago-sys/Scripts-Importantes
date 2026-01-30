import base64
import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# --- FUNÇÕES DE LÓGICA DE EXTRAÇÃO ---

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
    Lógica para extrair dados das Portarias. 
    Ajustada para identificar Metas, Itens e Valores.
    """
    items = []
    current_meta = ""
    
    # Regex para identificar padrões (Ex: Item 1, Meta 02, R$ 1.000,00)
    re_meta = re.compile(r"(?:Meta|META)\s*[:\-\s]*(\d+)", re.IGNORECASE)
    re_item = re.compile(r"(?:Item|ITEM)\s*[:\-\s]*(\d+)", re.IGNORECASE)
    re_valor = re.compile(r"R\$\s?(\d{1,3}(?:\.\d{3})*,\d{2})")

    for line in lines:
        # Tenta identificar a Meta atual
        meta_match = re_meta.search(line)
        if meta_match:
            current_meta = meta_match.group(1)
        
        # Tenta identificar um Item e Valor na linha
        item_match = re_item.search(line)
        valor_match = re_valor.search(line)
        
        if item_match:
            valor = valor_match.group(1) if valor_match else ""
            # Adiciona o dicionário com as chaves que sua planilha espera
            items.append({
                "Número da Meta Específica": current_meta,
                "Número do Item": item_match.group(1),
                "Descrição": line.strip(), # Pega a linha toda como descrição inicial
                "Valor Unitário": valor
            })
            
    return items 

def get_template_header_info(template_path):
    wb = load_workbook(template_path, data_only=True)
    ws = wb.active
    header_map = {}
    # Lê o cabeçalho na linha 2 (ajuste se a sua planilha for na linha 1)
    for cell in ws[2]: 
        if cell.value:
            header_map[str(cell.value).strip()] = cell.column
    return wb, header_map

def build_rows(parsed_items, header_map):
    rows = []
    for item in parsed_items:
        # Mapeia os dados extraídos para as colunas exatas do Excel
        row_data = {header: item.get(header, "") for header in header_map.keys()}
        rows.append(row_data)
    return rows

def generate_excel_bytes(template_path, rows, header_map):
    wb = load_workbook(template_path)
    ws = wb.active
    start_row = 3 # Começa a preencher na linha 3
    
    for i, row_data in enumerate(rows):
        for header, col_idx in header_map.items():
            valor = row_data.get(header, "")
            ws.cell(row=start_row + i, column=col_idx, value=valor)
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- INTERFACE STREAMLIT ---

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "Planilha Base.xlsx"
LOGO_PATH = BASE_DIR / "Logo.png"

st.set_page_config(page_title="Preenche Planilhas", page_icon="📄", layout="centered")

# CSS para estilo FAF
st.markdown("""
    <style>
    .header { display: flex; align-items: center; gap: 16px; }
    .header-title { font-size: 1.6rem !important; font-weight: 600; margin: 0; }
    .logo-wrap { width: 64px; height: 64px; border-radius: 16px; overflow: hidden; border: 1px solid #e6e6e6; }
    .logo-wrap img { width: 64px; height: 64px; object-fit: cover; }
    .brand-bar { display: grid; grid-template-columns: repeat(5, 1fr); height: 6px; border-radius: 999px; margin-top: 10px; margin-bottom: 20px; }
    .brand-bar span:nth-child(1) { background: #00b140; }
    .brand-bar span:nth-child(2) { background: #ff1b14; }
    .brand-bar span:nth-child(3) { background: #ffd200; }
    .brand-bar span:nth-child(4) { background: #1f4bff; }
    .brand-bar span:nth-child(5) { background: #ff1b14; }
    </style>
    """, unsafe_allow_html=True)

# Logo
logo_b64 = ""
if LOGO_PATH.exists():
    logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
logo_html = f'<div class="logo-wrap"><img src="data:image/png;base64,{logo_b64}" /></div>' if logo_b64 else ""

st.markdown(f'<div class="header">{logo_html}<h1 class="header-title">Gerador de Planilha de Itens - FAF</h1></div>', unsafe_allow_html=True)
st.markdown('<div class="brand-bar"><span></span><span></span><span></span><span></span><span></span></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("PDF do Plano de Aplicação", type=["pdf"])

if "result" not in st.session_state:
    st.session_state.result = None

if st.button("Processar Arquivos", type="primary", disabled=uploaded_file is None):
    if not TEMPLATE_PATH.exists():
        st.error("Erro: 'Planilha Base.xlsx' não encontrada no GitHub.")
    else:
        try:
            with st.status("Processando...", expanded=True) as status:
                lines = extract_lines_from_pdf_file(uploaded_file)
                parsed_items = parse_items(lines)
                
                if not parsed_items:
                    st.error("Nenhum item identificado no PDF. Verifique se o PDF tem texto selecionável.")
                else:
                    _, header_map = get_template_header_info(TEMPLATE_PATH)
                    rows = build_rows(parsed_items, header_map)
                    excel_bytes = generate_excel_bytes(TEMPLATE_PATH, rows, header_map)

                    st.session_state.result = {
                        "excel_bytes": excel_bytes,
                        "count": len(parsed_items)
                    }
                    status.update(label="Concluído!", state="complete")
        except Exception as e:
            st.error(f"Erro técnico: {e}")

if st.session_state.result:
    st.success(f"Sucesso! {st.session_state.result['count']} itens processados.")
    st.download_button(
        label="📥 Baixar Planilha Preenchida",
        data=st.session_state.result["excel_bytes"],
        file_name="Planilha_Portaria_Gerada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
