def build_rows(parsed_items, header_map):
    rows = []
    for item in parsed_items:
        row_data = {}
        for header in header_map.keys():
            h_lower = header.lower()
            val = ""
            
            # Mapeamento para preencher as colunas que você mencionou
            if "item" in h_lower and "número" in h_lower:
                val = item.get("Número do Item", "")
            elif "meta" in h_lower:
                # O PDF do CE costuma colocar a meta no campo Artigo ou labels similares
                val = item.get("Artigo", "") 
            elif "descrição" in h_lower:
                val = item.get("Descrição", "")
            elif "nome" in h_lower or "bem" in h_lower:
                val = item.get("Bem/Serviço", "")
            elif "unidade" in h_lower:
                val = item.get("Unidade de Medida", "")
            elif "quantidade" in h_lower:
                val = item.get("Qtd. Planejada", "")
            elif "valor" in h_lower and "total" in h_lower:
                val = item.get("Valor Total", "")
            elif "instituição" in h_lower or "órgão" in h_lower:
                val = item.get("Instituição", "")
            elif "status" in h_lower:
                val = "Aprovado" # Valor fixo para evitar o erro sequencial
            
            row_data[header] = val
        rows.append(row_data)
    return rows
