def parse_items(lines):
    items = []
    current_meta = None
    current_item = None
    current_status = None
    current_lines = []
    seen_items = set() # Trava para evitar repetição (como o item 30)

    def flush():
        nonlocal current_item, current_lines, current_status, current_meta
        if current_meta is None or current_item is None:
            return
        
        # Cria uma chave única para Meta + Item
        item_key = f"{current_meta}-{current_item}"
        if item_key not in seen_items:
            items.append({
                "meta": current_meta,
                "item": current_item,
                "status": current_status or "Planejado",
                "lines": current_lines[:],
            })
            seen_items.add(item_key)
        
        current_lines = []

    for line in lines:
        meta_match = META_RE.match(line)
        if meta_match:
            flush()
            current_meta = int(meta_match.group(1))
            continue
            
        item_match = ITEM_RE.match(line)
        if item_match:
            flush()
            current_item = int(item_match.group(1))
            current_status = (item_match.group(2) or "Planejado").capitalize()
            continue
            
        if current_item is not None:
            # Evita adicionar linhas de cabeçalho de página repetitivas
            if "Planos de Aplicação" in line or "CÓDIGO DE VERIFICAÇÃO" in line:
                continue
            current_lines.append(line)

    flush()
    return items

def extract_fields(item_lines):
    fields = {key: [] for key, _ in CAPTURE_PATTERNS}
    fields["acao"] = []
    fields["art"] = []
    fields["art_num"] = ""
    current_field = None

    for line in item_lines:
        # Tenta capturar o Artigo/Ação primeiro (é o que define a Meta na planilha)
        art_match = ART_PATTERN.match(line)
        if art_match:
            fields["art_num"] = art_match.group(1)
            fields["art"] = [art_match.group(3).strip()]
            current_field = "art"
            continue

        # Verifica se a linha pertence a um campo conhecido (Bem, Descrição, etc)
        matched_label = False
        for field, pattern in CAPTURE_PATTERNS:
            match = pattern.match(line)
            if match:
                content = match.group(1).strip()
                fields[field] = [content] if content else []
                current_field = field
                matched_label = True
                break
        
        if matched_label:
            continue

        # Se não houver novo rótulo, anexa a linha ao campo atual (acumula descrição longa)
        if current_field and line.strip():
            # Impede que metadados de sistema entrem nos campos
            if not any(stop.match(line) for stop in STOP_PATTERNS):
                fields[current_field].append(line.strip())

    # Consolida as listas em strings
    for key in fields:
        if isinstance(fields[key], list):
            fields[key] = " ".join(fields[key]).strip()
    
    return fields
