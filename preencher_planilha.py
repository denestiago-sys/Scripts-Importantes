import re
import io
import pandas as pd
from pypdf import PdfReader
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def extract_lines_from_pdf_file(file_searchable):
    """Extrai todas as linhas do PDF mantendo a ordem."""
    reader = PdfReader(file_searchable)
    all_lines = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            all_lines.extend(text.splitlines())
    return [line.strip() for line in all_lines if line.strip()]

def parse_items(lines):
    """
    Processa as linhas do PDF para extrair os itens sem limite de caracteres.
    Resolve o problema de repetição do Item 30 e campos em branco.
    """
    items = []
    current_item = None
    current_field = None
    
    # Rótulos que indicam o início de um novo dado
    labels = {
        "Art.": "Artigo",
        "Bem/Serviço:": "Bem/Serviço",
        "Descrição:": "Descrição",
        "Destinação:": "Destinação",
        "Cód. Senasp:": "Cód. Senasp",
        "Unidade de Medida:": "Unidade de Medida",
        "Qtd. Planejada:": "Qtd. Planejada",
        "Natureza (ND):": "Natureza (ND)",
        "Instituição:": "Instituição",
        "Valor Originário Planejado:": "Valor Originário",
        "Valor Total:": "Valor Total"
    }

    for line in lines:
        # 1. Identifica início de um novo Item (Ex: "Item 1", "Item 30")
        # O regex garante que a linha comece exatamente com "Item X"
        item_match = re.match(r"^Item\s+(\d+)$", line, re.IGNORECASE)
        
        if item_match:
            if current_item:
                items.append(current_item)
            current_item = {"Número do Item": item_match.group(1)}
            current_field = None
            continue

        if current_item is None:
            continue

        # 2. Verifica se a linha contém um dos rótulos principais
        found_label = False
        for label, key in labels.items():
            if label in line:
                # Extrai o valor após o ":"
                parts = line.split(label, 1)
                value = parts[1].strip() if len(parts) > 1 else ""
                current_item[key] = value
                current_field = key
                found_label = True
                break
        
        # 3. Se não tem rótulo e temos um campo ativo, é continuação (Texto Longo)
        if not found_label and current_field:
            # Ignora linhas de paginação ou cabeçalhos de sistema comuns no CE
            if "apps.mj.gov.br" in line or "Página" in line:
                continue
            
            # Acumula o texto sem limite
            prev_val = current_item.get(current_field, "")
            current_item[current_field] = f"{prev_val} {line}".strip()

    # Adiciona o último item processado
    if current_item:
        items.append(current_item)
        
    return items

def build_rows(parsed_items, header_map):
    """Mapeia os dados extraídos para as colunas da planilha base."""
    rows = []
    # Dicionário de tradução entre o que extraímos e o que a planilha espera
    mapping = {
        "Número do Item": ["Número do Item", "Item"],
        "Descrição": ["Descrição", "Descrição do Item"],
        "Bem/Serviço": ["Bem/Serviço", "Nome do Item"],
        "Destinação": ["Destinação", "Unidade Destinatária"],
        "Instituição": ["Instituição", "Órgão"],
        "Qtd. Planejada": ["Qtd. Planejada", "Quantidade"],
        "Valor Total": ["Valor Total", "Valor Estimado"]
    }

    for item in parsed_items:
        row_data = {}
        for header in header_map.keys():
            # Tenta encontrar o valor correspondente
            found_val = ""
            for internal_key, sheet_keys in mapping.items():
                if any(sk.lower() in header.lower() for sk in sheet_keys):
                    found_val = item.get(internal_key, "")
                    break
            row_data[header] = found_val
        rows.append(row_data)
    return rows

def generate_excel_bytes(template_path, rows, header_map, **kwargs):
    """Gera o arquivo Excel final com os dados acumulados."""
    wb = load_workbook(template_path)
    ws = wb.active
    
    # Inicia o preenchimento a partir da linha 3 (ajuste se necessário)
    start_row = 3
    
    for r_idx, row_data in enumerate(rows, start=start_row):
        for header, col_idx in header_map.items():
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.value = row_data.get(header)
            # Garante que o Excel aceite textos longos com quebra de linha
            cell.alignment = cell.alignment.copy(wrapText=True)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# Funções auxiliares (stubs para compatibilidade com seu Streamlit)
def is_analysis_template_file(path): return False
def get_template_header_info(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    header_map = {}
    # Lê a linha 2 para mapear cabeçalhos (ajuste se sua planilha for diferente)
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=2, column=col).value
        if val: header_map[str(val)] = col
    return 2, header_map

def extract_plan_signature(lines): return {"sigla": "CE", "ano": "2025"}
def resolve_art_by_plan_rule(sigla, ano): return "N/A"
