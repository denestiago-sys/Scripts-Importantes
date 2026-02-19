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
            lines = [l.strip() for l in text.splitlines() if "apps.mj.gov.br" not in l and "Página" not in l]
            all_lines.extend(lines)
    return [line for line in all_lines if line]

def parse_items(lines):
    items = []
    current_item = None
    current_field = None
    labels = {
        "Bem/Serviço:": "Bem/Serviço",
        "Descrição:": "Descrição",
        "Destinação:": "Destinação",
        "Unidade de Medida:": "Unidade de Medida",
        "Qtd. Planejada:": "Qtd. Planejada",
        "Instituição:": "Instituição",
        "Valor Total:": "Valor Total"
    }

    for line in lines:
        clean_line = line.strip()
        # Detecta "Item X" apenas no início da linha para evitar erro do "Item 30"
        item_match = re.match(r"^(?:Item|ITEM)\s+(\d+)", clean_line)
        
        if item_match:
            if current_item: items.append(current_item)
            current_item = {"Número do Item": item_match.group(1), "Artigo": ""}
            current_field = None
            # Captura a Meta/Artigo se estiver na mesma linha do Item
            if "Art." in clean_line:
                current_item["Artigo"] = clean_line.split("Item")[0].strip() or clean_line.split(item_match.group(0))[-1].strip()
            continue

        if current_item is None: continue

        # Captura de campos rotulados
        found_label = False
        for label, key in labels.items():
            if label in clean_line:
                current_item[key] = clean_line.split(label, 1)[1].strip()
                current_field = key
                found_label = True
                break
        
        # Acúmulo de texto para campos longos (Descrição) e Meta (Art.)
        if not found_label:
            if "Art." in clean_line:
                current_item["Artigo"] = (current_item.get("Artigo", "") + " " + clean_line).strip()
            elif current_field and ":" not in clean_line:
                current_item[current_field] = (current_item.get(current_field, "") + " " + clean_line).strip()

    if current_item: items.append(current_item)
    return items

def build_rows(parsed_items, header_map):
    rows = []
    for item in parsed_items:
        row_data = {}
        for header in header_map.keys():
            h_lower = header.lower()
            val = ""
            if "item" in h_lower and "número" in h_lower: val = item.get("Número do Item", "")
            elif "meta" in h_lower: val = item.get("Artigo", "")
            elif "descrição" in h_lower: val = item.get("Descrição", "")
            elif "nome" in h_lower or "bem" in h_lower: val = item.get("Bem/Serviço", "")
            elif "unidade" in h_lower: val = item.get("Unidade de Medida", "")
            elif "quantidade" in h_lower: val = item.get("Qtd. Planejada", "")
            elif "valor" in h_lower: val = item.get("Valor Total", "")
            elif "órgão" in h_lower: val = item.get("Instituição", "")
            elif "status" in h_lower: val = "Aprovado"
            row_data[header] = val
        rows.append(row_data)
    return rows

def generate_excel_bytes(template_path, rows, header_map):
    wb = load_workbook(template_path)
    ws = wb.active
    for r_idx, row_data in enumerate(rows, start=3):
        for header, col_idx in header_map.items():
            cell = ws.cell(row=r_idx, column=col_idx)
            val = row_data.get(header, "")
            # Trata valores numéricos para o Excel
            if "valor" in header.lower() and val:
                try:
                    num = float(str(val).replace("R$", "").replace(".", "").replace(",", ".").strip())
                    cell.value = num
                    cell.number_format = '#,##0.00'
                except: cell.value = val
            else:
                cell.value = val
            cell.alignment = cell.alignment.copy(wrapText=True, vertical='top')
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def get_template_header_info(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    header_map = {str(ws.cell(row=2, column=c).value): c for c in range(1, ws.max_column + 1) if ws.cell(row=2, column=c).value}
    return 2, header_map
