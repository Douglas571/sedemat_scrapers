import re
import csv
import json

def extract_information(markdown_text):
    # Remove all markdown bold asterisks to simplify parsing
    clean_text = markdown_text.replace('**', '')
    
    # Split the text into blocks for each "N° DE PATENTE"
    blocks = re.split(r'N[°º]\s*DE\s*PATENTE:?\s*', clean_text)
    
    extracted_data = []
    
    # Dictionary to map Spanish months to numbers (includes typo fix for 'FBRERO')
    meses = {
        'ENERO': '01', 'FEBRERO': '02', 'FBRERO': '02', 'MARZO': '03', 
        'ABRIL': '04', 'MAYO': '05', 'JUNIO': '06', 'JULIO': '07', 
        'AGOSTO': '08', 'SEPTIEMBRE': '09', 'OCTUBRE': '10', 
        'NOVIEMBRE': '11', 'DICIEMBRE': '12'
    }
    
    for block in blocks:
        if not block.strip():
            continue
            
        # Initialize a dictionary for the current license
        license_data = {}
        
        # 1. Patente Number
        license_data['N_PATENTE'] = block.split('\n')[0].strip()
        
        # Code (Extracts the middle number between slashes, e.g., 001 from SAT/001/2026)
        code_match = re.search(r'/(\d+)/', license_data['N_PATENTE'])
        license_data['CODIGO'] = code_match.group(1) if code_match else None
        
        # 2. Company Name
        name_match = re.search(r'QUE SE OTORGA(?:\s+A:?)?\s*\n+([^\n]+)', block)
        license_data['EMPRESA'] = name_match.group(1).strip() if name_match else None
        
        # 3. R.I.F / C.I.
        rif_match = re.search(r'C\.I\./R\.I\.F\.?:?\s*([VJGVE]-\d+-\d+|\d+)', block, re.IGNORECASE)
        license_data['RIF'] = rif_match.group(1).strip() if rif_match else None
        
        # 4. Dirección
        dir_match = re.search(r'DIRECCI[OÓ]N(?: FISCA)?:?\s*([^\n]+)', block, re.IGNORECASE)
        license_data['DIRECCION'] = dir_match.group(1).strip() if dir_match else None
        
        # 5. Ramo
        ramo_match = re.search(r'EN EL RAMO DE:?\s*([^\n]+)', block, re.IGNORECASE)
        license_data['RAMO'] = ramo_match.group(1).strip() if ramo_match else None
        
        # 6. Declaración
        dec_match = re.search(r'DECLARACI[OÓ]N N[°º]:?\s*([^\n]+)', block, re.IGNORECASE)
        license_data['DECLARACION'] = dec_match.group(1).strip() if dec_match else None
        
        # 7. Representante Legal
        rep_match = re.search(r'REPRESENTANTE LEGAL:?\s*([^\n]+)', block, re.IGNORECASE)
        license_data['REPRESENTANTE_LEGAL'] = rep_match.group(1).strip() if rep_match else None
        
        # 8. Cédula de Identidad
        cedula_match = re.search(r'C[EÉ]DULA DE IDENTIDAD:?\s*([^\n]+)', block, re.IGNORECASE)
        license_data['CEDULA_REPRESENTANTE'] = cedula_match.group(1).strip() if cedula_match else None
        
        # 9. IMPROVEMENT: Split Ramo Aforo into Code and Value
        aforo_match = re.search(r'RAMO AFORO\s*\n+([^\n]+)', block)
        if aforo_match:
            aforo_line = aforo_match.group(1).strip()
            aforo_parts = aforo_line.split()  # Splits by whitespace
            
            # Ensure there are at least two items to unpack securely
            license_data['RAMO_AFORO_CODIGO'] = aforo_parts[0] if len(aforo_parts) > 0 else None
            license_data['RAMO_AFORO_VALOR'] = aforo_parts[1] if len(aforo_parts) > 1 else None
        else:
            license_data['RAMO_AFORO_CODIGO'] = None
            license_data['RAMO_AFORO_VALOR'] = None
        
        # 10. Date / Location (e.g., PUERTO CUMAREBO, 16 DE ENERO DEL 2026.)
        date_match = re.search(r'PUERTO CUMAREBO[,;]\s*([^\n]+202\d\.?)', block, re.IGNORECASE)
        if date_match:
            raw_date = date_match.group(1).strip(' .')
            license_data['FECHA_TEXTO'] = raw_date
            
            # Standardized Date (YYYY-MM-DD)
            parts_match = re.search(r'(\d{1,2})\s+DE\s+([A-Z]+)\s+DEL?\s+(\d{4})', raw_date, re.IGNORECASE)
            if parts_match:
                day = parts_match.group(1).zfill(2)
                month_str = parts_match.group(2).upper()
                year = parts_match.group(3)
                
                month = meses.get(month_str, '01')
                license_data['FECHA_ISO'] = f"{year}-{month}-{day}"
            else:
                license_data['FECHA_ISO'] = None
        else:
            license_data['FECHA_TEXTO'] = None
            license_data['FECHA_ISO'] = None
            
        extracted_data.append(license_data)
        
    return extracted_data

def save_to_csv(data, filename="patentes_extraidas.csv"):
    if not data:
        print("No data to save.")
        return
        
    headers = data[0].keys()
    
    with open(filename, mode='w', encoding='utf-8-sig', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=headers)
        writer.writeheader()
        writer.writerows(data)
    print(f"Data successfully saved to {filename}")

if __name__ == "__main__":
    try:
        with open('patentes.md', 'r', encoding='utf-8') as f:
            markdown_text = f.read()
            
        parsed_data = extract_information(markdown_text)
        
        print("Sample extracted record:")
        print(json.dumps(parsed_data[0], indent=4, ensure_ascii=False))
        print("-" * 40)
        
        save_to_csv(parsed_data)
        
    except FileNotFoundError:
        print("Please save your text in a file named 'patentes.md' in the same directory.")