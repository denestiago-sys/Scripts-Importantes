import re
import io
import pandas as pd
from pypdf import PdfReader
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def extract_lines_from_pdf_file(file_searchable):
    """Extrai linhas de forma limpa, tratando quebras de página do CE."""
    reader = PdfReader(file_searchable)
    all_lines = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            # Filtra rodapés e metadados que cortam a descrição
            lines = [l.strip() for l in text.splitlines() if "apps.mj.gov.br" not in l and "Página" not in l]
            all_lines.extend(lines)
    return [line for line in all_lines if line]

def parse_items(lines):
    """Lógica avançada para detectar itens e ignorar citações como 'Item 30'."""
    items = []
    current_item = None
    current_field = None
    
    # Rótulos padrão do sistema do Ministério da Justiça
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
        clean_line = line.strip()
        
        # 1. DETECÇÃO DE NOVO ITEM (O SEGREDO ESTÁ AQUI)
        # Procuramos por "Item X" onde X é um número. 
        # Para evitar o erro do "Item 30", verificamos se a linha contém APENAS isso 
        # ou se a palavra "Item" aparece logo após uma quebra de contexto.
        is_new_item = False
        item_match = re.search(r"^(?:Item|ITEM)\s+(\d+)$", clean_line)
        
        if item_match:
            is_new_item = True
        # Caso o PDF junte o Item com o Artigo na mesma linha (comum no CE)
        elif clean_line.startswith("Item ") and "Art." in clean_line:
            item_match = re.search(r"Item\s+(\d+)", clean_line)
            is_new_item = True

        if is_new_item and item_match:
            if current_item:
                items.append(current_item)
            current_item = {"Número do Item": item_match.group(1)}
            current_field = None
            
            # Se a linha tinha mais coisa (ex: Artigo), tentamos processar o resto
            remainder = clean_line.replace(item_match.group(0), "").strip()
            if "Art." in remainder:
                current_item["Artigo"] = remainder
            continue

        if current_item is None:
            continue

        # 2. CAPTURA DE CAMPOS COM RÓTULO
        found_label = False
        for label, key in labels.items():
            if label in clean_line:
                parts = clean_line.split(label, 1)
                current_item[key] = parts[1].strip() if len(parts) > 1 else ""
                current_field = key
                found_label = True
                break
        
        # 3. ACÚMULO DE TEXTO (RESOLVE CAMPOS EM BRANCO E TEXTOS LONGOS)
        if not found_label and current_field:
            # Se a linha NÃO contém ':' (o que indica um novo campo)
            # E não é uma citação de item que o regex acima já filtraria
            if ":" not in clean_line:
                prev_val = current_item.get(current_field, "")
                # Concatena sem limite de caracteres
                current_item[current_field] = f"{prev_val} {clean_line}".strip()

    if current_item:
        items.append(current_item)
    return items

def build_rows(parsed_items, header_map):
    """Mapeia os dados extraídos para as colunas da planilha."""
    rows = []
    mapping = {
        "Número do Item": ["item", "nº"],
        "Descrição": ["descrição"],
        "Bem/Serviço": ["bem", "serviço", "nome"],
        "Destinação": ["destinação", "unidade destinatária"],
        "Instituição": ["instituição", "órgão"],
        "Qtd. Planejada": ["qtd", "quantidade"],
        "Valor Total": ["valor total", "estimado"]
    }

    for item in parsed_items:
        row_data = {}
        for header in header_map.keys():
            val = ""
            header_lower = header.lower()
            for internal_key, keywords in mapping.items():
                if any(kw in header_lower for kw in keywords):
                    val = item.get(internal_key, "")
                    break
            row_data[header] = val
        rows.append(row_data)
    return rows

def generate_excel_bytes(template_path, rows, header_map, **kwargs):
    wb = load_workbook(template_path)
    ws = wb.active
    
    # Preenchimento a partir da linha 3
    start_row = 3
    for r_idx, row_data in enumerate(rows, start=start_row):
        for header, col_idx in header_map.items():
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.value = row_data.get(header)
            cell.alignment = cell.alignment.copy(wrapText=True) # Ativa quebra de texto

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- FUNÇÕES DE COMPATIBILIDADE (STUBS) ---
def get_template_header_info(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    header_map = {}
    # Lê a linha 2 da Planilha Base para saber as colunas
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
