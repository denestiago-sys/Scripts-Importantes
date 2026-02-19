import re
import openpyxl
import pdfplumber
import io
from io import BytesIO
from pathlib import Path

# --- PADRÕES DE BUSCA (REGEX) ---
META_RE = re.compile(r"^META ESPEC[ÍI]FICA\s+(\d+)", re.IGNORECASE)
ITEM_RE = re.compile(r"^Item\s*(\d+)\s*(Planejado|Aprovado|Cancelado)?", re.IGNORECASE)
ART_PATTERN = re.compile(r"^Art\.?\s*(6|7|8)\s*º?\s*(?:\((\d+)\))?\s*:\s*(.*)", re.IGNORECASE)

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

def normalize(text):
    return re.sub(r"\s+", " ", text or "").strip()

def extract_lines_from_pdf(file_obj):
    file_obj.seek(0)
    lines = []
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            text = text.replace("\x0c", "\n")
            for raw in text.splitlines():
                if raw.strip():
                    lines.append(raw.strip())
    return lines

def parse_items(lines):
    items = []
    current_meta = "N/A"
    current_item_obj = None
    seen_items = set() 

    for line in lines:
        m_meta = META_RE.match(line)
        if m_meta:
            current_meta = m_meta.group(1)
            continue
        
        m_item = ITEM_RE.match(line)
        if m_item:
            item_num = m_item.group(1)
            unique_key = f"M{current_meta}I{item_num}" 
            
            if unique_key not in seen_items:
                current_item_obj = {
                    "meta_num": current_meta,
                    "item_num": item_num,
                    "status": m_item.group(2) or "Planejado",
                    "lines": []
                }
                items.append(current_item_obj)
                seen_items.add(unique_key)
            continue
        
        if current_item_obj:
            current_item_obj["lines"].append(line)
    return items

def extract_fields(item_lines):
    fields = {key: [] for key, _ in CAPTURE_PATTERNS}
    fields.update({"art_texto": ""})
    current_f = None
    
    for line in item_lines:
        m_art = ART_PATTERN.match(line)
        if m_art:
            fields["art_texto"] = line # Guarda o texto longo do Artigo
            continue

        matched = False
        for key, pat in CAPTURE_PATTERNS:
            m = pat.match(line)
            if m:
                fields[key].append(m.group(1))
                current_f = key
                matched = True
                break
        
        if not matched and current_f:
            fields[current_f].append(line)
            
    return {k: normalize(" ".join(v)) if isinstance(v, list) else v for k, v in fields.items()}

def build_rows(parsed_items):
    rows = []
    for item in parsed_items:
        f = extract_fields(item["lines"])
        rows.append({
            "Número da Meta Específica": item["meta_num"], # APENAS O NÚMERO
            "Número do Item": item["item_num"],
            "Ação conforme Art. 7º da portaria nº 685": f["art_texto"], # TEXTO LONGO DO ARTIGO
            "Material/Serviço": f["bem"],
            "Descrição": f["descricao"],
            "Destinação": f["destinacao"],
            "Instituição": f["instituicao"],
            "Natureza da Despesa": f["natureza"],
            "Quantidade Planejada": f["quantidade"],
            "Unidade de Medida": f["unidade"],
            "Valor Planejado Total": f["valor_total"],
            "Status do Item": item["status"]
        })
    return rows

def generate_excel_bytes(template_path, rows):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    header_row = 1
    # Procura a linha que contém os títulos das colunas
    for r in range(1, 10):
        val = str(ws.cell(r, 1).value or "")
        if "Meta" in val or "Número" in val:
            header_row = r
            break

    # Mapeia os nomes das colunas para os índices (1, 2, 3...)
    col_map = {str(ws.cell(header_row, c).value).strip(): c 
               for c in range(1, ws.max_column + 1) if ws.cell(header_row, c).value}

    for idx, data in enumerate(rows, start=header_row + 1):
        for header_name, col_idx in col_map.items():
            # Tenta encontrar a chave correspondente no dicionário de dados
            if header_name in data:
                ws.cell(row=idx, column=col_idx, value=data[header_name])

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
