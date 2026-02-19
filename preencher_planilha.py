def build_rows(parsed_items, header_map):
    rows = []
    
    # Mapeamento refinado para o padrão do Ceará / Ministério da Justiça
    for item in parsed_items:
        row_data = {}
        for header in header_map.keys():
            h_lower = header.lower()
            val = ""
            
            if "item" in h_lower and "número" in h_lower:
                val = item.get("Número do Item", "")
            elif "meta" in h_lower:
                # Tenta pegar do campo Artigo que costuma conter o número da meta no CE
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
            elif "órgão" in h_lower or "instituição" in h_lower:
                val = item.get("Instituição", "")
            elif "destinação" in h_lower or "unidade destinatária" in h_lower:
                val = item.get("Destinação", "")
            elif "status" in h_lower:
                # Pega o status real se houver, ou deixa vazio em vez de número sequencial
                val = item.get("Status", "Aprovado") 
                
            row_data[header] = val
        rows.append(row_data)
    return rows

def generate_excel_bytes(template_path, rows, header_map, **kwargs):
    wb = load_workbook(template_path)
    ws = wb.active
    
    # Detecta automaticamente onde começam os dados se não for na linha 3
    start_row = 3 
    
    for r_idx, row_data in enumerate(rows, start=start_row):
        for header, col_idx in header_map.items():
            cell = ws.cell(row=r_idx, column=col_idx)
            
            # Limpeza especial para valores financeiros (Valor Planejado Total)
            valor = row_data.get(header, "")
            if "Valor" in header and isinstance(valor, str):
                # Remove "R$", pontos de milhar e troca vírgula por ponto para o Excel entender como número
                limpo = valor.replace("R$", "").replace(".", "").replace(",", ".").strip()
                try:
                    cell.value = float(limpo)
                    cell.number_format = '#,##0.00'
                except:
                    cell.value = valor
            else:
                cell.value = valor
                
            cell.alignment = cell.alignment.copy(wrapText=True, vertical='top')

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
