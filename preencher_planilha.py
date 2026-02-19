import argparse
import copy
import re
import io
import pdfplumber
import openpyxl
from io import BytesIO
from pathlib import Path

# --- CONFIGURAÇÕES ORIGINAIS PRESERVADAS ---
META_RE = re.compile(r"^META ESPEC[ÍI]FICA\s+(\d+)", re.IGNORECASE)
ITEM_RE = re.compile(r"^Item\s*(\d+)\s*(Planejado|Aprovado|Cancelado)?", re.IGNORECASE)
ACTION_HEADER_KEY = "acao_art"
ACTION_HEADER_NUM_KEY = "acao_art_num"
ACTION_HEADER_PATTERN = re.compile(r"^Ação conforme Art\.\s*\d+º\s+da portaria nº 685$", re.IGNORECASE)
PLAN_SIGNATURE_RE = re.compile(r"\b([A-Z]{2})\s*-\s*([A-Z0-9]+)\s*-\s*(20\d{2})\b")
ART_PATTERN = re.compile(r"^Art\.?\s*(6|7|8)\s*º?\s*(?:\((\d+)\))?\s*:\s*(.*)", re.IGNORECASE)
ACTION_PATTERN = re.compile(r"^A[cç][aã]o:\s*(.*)", re.IGNORECASE)

CAPTURE_PATTERNS = [
    ("bem", re.compile(r"^(?:Bem|Material)/Servi[cç]o:\s*(.*)", re.IGNORECASE)),
    ("descricao", re.compile(r"^Descri[cç][aã]o:\s*(.*)", re.IGNORECASE)),
    ("destinacao", re.compile(r"^Destina[cç][aã]o:\s*(.*)", re.IGNORECASE)),
    ("unidade", re.compile(r"^Unidade de Medida:\s*(.*)", re.IGNORECASE)),
    ("quantidade", re.compile(r"^Qtd\.?\s*Planejada:\s*(.*)", re.IGNORECASE)),
    ("natureza", re.compile(r"^Natureza\s*\(ND\):\s*(.*)", re.IGNORECASE)),
    ("instituicao", re.compile(r"^Institui[cç][aã]o:\s*(.*)", re.IGNORECASE)),
    ("valor_total", re.compile(r"^Valor Total:\s*(.*)", re.IGNORECASE)),
]

STOP_PATTERNS = [re.compile(r"^C[oó]d\.?\s*Senasp:", re.IGNORECASE), re.compile(r"^Valor Origin[aá]rio Planejado:", re.IGNORECASE)]

# --- FUNÇÕES DE UTILIDADE ---
def normalize(text):
    return re.sub(r"\s+", " ", text or "").strip()

def format_currency(value):
    digits = re.sub(r"[^0-9,]", "", value or "")
    if not digits: return ""
    return f"R$ {digits}"

def parse_int(value):
    digits = re.sub(r"[^0-9]", "", value or "")
    return int(digits) if digits else ""

def normalize_pdf_text(text):
    text = text.replace("\x0c", "\n")
    text = re.sub(r"(META ESPEC[ÍI]FICA\s+\d+)", r"\n\1\n", text, flags=re.IGNORECASE)
    text = re.sub(r"(Item\s*\d+)", r"\n\1\n", text, flags=re.IGNORECASE)
    return text

def extract_lines_from_pdf_file(file_obj):
    file_obj.seek(0)
    lines = []
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            text = normalize_pdf_text(page.extract_text() or "")
            for raw in text.splitlines():
                if raw.strip() and "apps.mj.gov.br" not in raw:
                    lines.append(raw.strip())
    return lines

def parse_items(lines):
    items = []
    current_meta = ""
    current_item = None
    for line in lines:
        m_meta = META_RE.match(line)
        if m_meta:
            current_meta = m_meta.group(1)
            continue
        m_item = ITEM_RE.match(line)
        if m_item:
            if current_item: items.append(current_item)
            current_item = {"meta": current_meta, "item": m_item.group(1), "status": (m_item.group(2) or "Aprovado"), "lines": []}
            continue
        if current_item: current_item["lines"].append(line)
    if current_item: items.append(current_item)
    return items

def extract_fields(item_lines):
    fields = {key: [] for key, _ in CAPTURE_PATTERNS}
    fields.update({"acao": [], "art": [], "art_num": ""})
    current_f = None
    for line in item_lines:
        matched = False
        for key, pat in CAPTURE_PATTERNS:
            m = pat.match(line)
            if m:
                fields[key].append(m.group(1))
                current_f = key
                matched = True; break
        if not matched:
            m_art = ART_PATTERN.match(line)
            if m_art:
                fields["art"].append(m_art.group(3))
                fields["art_num"] = m_art.group(1)
                current_f = "art"; matched = True
        if not matched and current_f:
            fields[current_f].append(line)
    return {k: normalize(" ".join(v)) if isinstance(v, list) else v for k, v in fields.items()}

def build_rows(parsed_items, header_map):
    rows = []
    for item in parsed_items:
        f = extract_fields(item["lines"])
        row = {
            "Número da Meta Específica": item["meta"],
            "Número do Item": item["item"],
            "Ação conforme Art. 7º da portaria nº 685": f["art"] or f["acao"],
            "Material/Serviço": f["bem"],
            "Descrição": f["descricao"],
            "Destinação": f["destinacao"],
            "Instituição": f["instituicao"],
            "Natureza da Despesa": f["natureza"],
            "Quantidade Planejada": f["quantidade"],
            "Unidade de Medida": f["unidade"],
            "Valor Planejado Total": format_currency(f["valor_total"]),
            "Status do Item": item["status"],
        }
        rows.append(row)
    return rows

def get_template_header_info(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    # Procura a linha de cabeçalho (geralmente a 2 ou onde houver "Item")
    header_row = 2
    for r in range(1, 5):
        if "item" in str(ws.cell(r, 1).value or "").lower():
            header_row = r; break
    header_map = {str(ws.cell(header_row, c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(header_row, c).value}
    return header_row, header_map

def generate_excel_bytes(template_path, rows, header_map):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    start_row = 3 # Ajuste conforme seu modelo
    for idx, row_data in enumerate(rows, start=start_row):
        for header, col_idx in header_map.items():
            val = row_data.get(header, "")
            # Mapeamento flexível de nomes de colunas
            if not val:
                for k, v in row_data.items():
                    if header.lower() in k.lower() or k.lower() in header.lower():
                        val = v; break
            ws.cell(row=idx, column=col_idx, value=val)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
