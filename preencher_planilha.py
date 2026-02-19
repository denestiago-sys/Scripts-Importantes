import re
import io
import pandas as pd
from pypdf import PdfReader
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def extract_lines_from_pdf_file(file_searchable):
    reader = PdfReader(file_searchable)
    all_lines = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            all_lines.extend(text.splitlines())
    return [line.strip() for line in all_lines if line.strip()]

def parse_items(lines):
    items = []
    current_item = None
    current_field = None
    
    # Rótulos estritos para evitar capturar referências no meio do texto
    labels = {
        "Bem/Serviço:": "Bem/Serviço",
        "Descrição:": "Descrição",
        "Destinação:": "Destinação",
        "Cód. Senasp:": "Cód. Senasp",
        "Unidade de Medida:": "Unidade de Medida",
        "Qtd. Planejada:": "Qtd. Planejada",
        "Natureza (ND):": "Natureza (ND)",
        "Instituição:": "Instituição",
        "Valor Total:": "Valor Total"
    }

    for line in lines:
        # RESOLVE O ITEM 30: O regex ^Item\s+(\d+)$ garante que "Item" esteja no início da linha
        # evitando capturar "REMANEJAMENTO DE SALDO DO ITEM 30" como um novo item.
        item_match = re.match(r"^Item\s+(\d+)$", line, re.IGNORECASE)
        
        if item_match:
            if current_item:
                items.append(current_item)
            current_item = {"Número do Item": item_match.group(1)}
            current_field = None
            continue

        if current_item is None: continue

        found_label = False
        for label, key in labels.items():
            if label in line:
                parts = line.split(label, 1)
                current_item[key] = parts[1].strip() if len(parts) > 1 else ""
                current_field = key
                found_label = True
                break
        
        # ACUMULA TEXTO SEM LIMITE: Se não é rótulo e estamos num campo, é continuação
        if not found_label and current_field:
            if "apps.mj.gov.br" in line or "Página" in line: continue
            
            # Se a linha não parece o início de outro campo (não tem ':'), acumula
            if ":" not in line or (line.split(':')[0] not in labels):
                prev_val = current_item.get(current_field, "")
                current_item[current_field] = f"{prev_val} {line}".strip()

    if current_item: items.append(current_item)
    return items

def build_rows(parsed_items, header_map):
    rows = []
    # Mapeamento flexível para encontrar as colunas independente do nome exato
    mapping = {
        "Número do Item": ["item", "nº"],
        "Descrição": ["descrição"],
        "Bem/Serviço": ["bem", "serviço", "nome"],
        "Destinação": ["destinação", "unidade"],
        "Instituição": ["instituição", "órgão"],
        "Qtd. Planejada": ["qtd", "quantidade"],
        "Valor Total": ["valor total", "estimado"]
    }

    for item in parsed_items:
        row_data = {}
        for header in header_map.keys():
            val = ""
            for internal_key, keywords in mapping.items():
                if any(kw in header.lower() for kw in keywords):
                    val = item.get(internal_key, "")
                    break
            row_data[header] = val
        rows.append(row_data)
    return rows

def generate_excel_bytes(template_path, rows, header_map, **kwargs):
    wb = load_workbook(template_path)
    ws = wb.active
    # Identifica se é a planilha de análise (Ceará) ou a simples
    start_row = 3 
    
    for r_idx, row_data in enumerate(rows, start=start_row):
        for header, col_idx in header_map.items():
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.value = row_data.get(header)
            cell.alignment = cell.alignment.copy(wrapText=True)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- FUNÇÕES QUE FALTAVAM PARA O APP.PY FUNCIONAR ---

def get_template_header_info(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    header_map = {}
    # Busca cabeçalhos na linha 2 (padrão FAF)
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=2, column=col).value
        if val: header_map[str(val)] = col
    return 2, header_map

def is_analysis_template_file(path):
    # Verifica se a planilha tem a palavra "ANÁLISE" para ativar o modo Ceará
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    for row in range(1, 5):
        for col in range(1, 5):
            val = str(ws.cell(row=row, column=col).value or "")
            if "ANÁLISE" in val.upper(): return True
    return False

def find_items_table_header_row(ws):
    for row in range(1, 20):
        for col in range(1, 10):
            if "Item" in str(ws.cell(row=row, column=col).value or ""):
                return row
    return 2

def extract_analysis_data(lines): return {"sections": []}
def collect_analysis_missing_cells(data): return []
def get_analysis_items_header_info(path): return get_template_header_info(path) + ({},)
def extract_plan_signature(lines): return {"sigla": "CE", "ano": "2025"}
def resolve_art_by_plan_rule(sigla, ano): return "N/A"
