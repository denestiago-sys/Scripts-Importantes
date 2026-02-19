import re
import io
import pandas as pd
from pypdf import PdfReader
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def extract_lines_from_pdf_file(file_searchable):
    """Extrai todas as linhas do PDF removendo ruídos de cabeçalho."""
    reader = PdfReader(file_searchable)
    all_lines = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            # Remove linhas que são links ou números de página para não quebrar o texto
            lines = [l.strip() for l in text.splitlines() if "apps.mj.gov.br" not in l and "Página" not in l]
            all_lines.extend(lines)
    return [line for line in all_lines if line]

def parse_items(lines):
    """Extrai itens garantindo que citações (como Item 30) não criem novos itens."""
    items = []
    current_item = None
    current_field = None
    
    # Rótulos para busca no PDF do Ceará
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
        # REGEX MILIMÉTRICO: Só aceita "Item X" se for a única coisa na linha ou estiver isolado.
        # Isso ignora "REMANEJAMENTO DE SALDO DO ITEM 30"
        item_match = re.match(r"^(?:Item|ITEM)\s+(\d+)$", line.strip())
        
        if item_match:
            if current_item:
                items.append(current_item)
            current_item = {"Número do Item": item_match.group(1)}
            current_field = None
            continue

        if current_item is None:
            continue

        found_label = False
        for label, key in labels.items():
            if label in line:
                parts = line.split(label, 1)
                current_item[key] = parts[1].strip() if len(parts) > 1 else ""
                current_field = key
                found_label = True
                break
        
        # CONTINUAÇÃO DE TEXTO: Se não tem rótulo e não é um novo item, concatena ao campo anterior
        if not found_label and current_field:
            # Se a linha não parece o início de outro metadado (não tem ':'), acumula tudo
            if ":" not in line:
                prev_val = current_item.get(current_field, "")
                current_item[current_field] = f"{prev_val} {line}".strip()

    if current_item:
        items.append(current_item)
    return items

def build_rows(parsed_items, header_map):
    """Faz o De-Para entre os campos do PDF e as colunas da sua Planilha Base."""
    rows = []
    mapping = {
        "Número do Item": ["item"],
        "Descrição": ["descrição"],
        "Bem/Serviço": ["bem", "serviço", "nome"],
        "Destinação": ["destinação", "unidade destinatária"],
        "Instituição": ["instituição", "órgão"],
        "Qtd. Planejada": ["qtd", "quantidade"],
        "Valor Total": ["valor total"]
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
    
    # Começa na linha 3 (abaixo do cabeçalho da linha 2)
    start_row = 3
    
    for r_idx, row_data in enumerate(rows, start=start_row):
        for header, col_idx in header_map.items():
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.value = row_data.get(header)
            # Permite que o Excel quebre o texto para exibir descrições longas
            cell.alignment = cell.alignment.copy(wrapText=True)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- STUBS PARA COMPATIBILIDADE COM SEU APP.PY ---
def get_template_header_info(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    header_map = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=2, column=col).value
        if val: header_map[str(val)] = col
    return 2, header_map

def is_analysis_template_file(path): return False
def extract_analysis_data(lines): return {"sections": []}
def collect_analysis_missing_cells(data): return []
def get_analysis_items_header_info(path): return get_template_header_info(path) + ({},)
def find_items_table_header_row(ws): return 2
def extract_plan_signature(lines): return {"sigla": "CE", "ano": "2025"}
def resolve_art_by_plan_rule(sigla, ano): return "N/A"
