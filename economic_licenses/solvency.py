import re
import csv
import json

def extraer_solvencias(markdown_text):
    # Remove all markdown bold marks
    clean_text = markdown_text.replace('**', '')
    
    # FIX: Split only when SAT-AE is located at the absolute start of a line
    bloques = re.split(r'(?:\r?\n|^)(?=SAT-AE/\d+/)', clean_text)
    
    lista_solvencias = []
    meses = {
        'ENERO': '01', 'FEBRERO': '02', 'FBRERO': '02', 'MARZO': '03', 
        'ABRIL': '04', 'MAYO': '05', 'JUNIO': '06', 'JULIO': '07', 
        'AGOSTO': '08', 'SEPTIEMBRE': '09', 'OCTUBRE': '10', 
        'NOVIEMBRE': '11', 'DICIEMBRE': '12'
    }
    
    for bloque in bloques:
        if not bloque.strip() or "CERTIFICADO DE SOLVENCIA" not in bloque:
            continue
            
        solvencia_data = {}
        lines = [line.strip() for line in bloque.splitlines() if line.strip()]
        
        # 1. Certificado & Código
        solvencia_data['N_CERTIFICADO'] = lines[0]
        codigo_match = re.search(r'/(\d+)/', lines[0])
        solvencia_data['CODIGO'] = codigo_match.group(1) if codigo_match else None
        
        # 2. Contribuyente / Empresa (Line right after the main text block)
        for i, line in enumerate(lines):
            if "CERTIFICA QUE EL CONTRIBUYENTE" in line:
                if i + 1 < len(lines):
                    solvencia_data['CONTRIBUYENTE'] = lines[i+1]
                break
        if 'CONTRIBUYENTE' not in solvencia_data:
            solvencia_data['CONTRIBUYENTE'] = None
            
        # 3. Cédula o RIF
        rif_match = re.search(r'CÉDULA/RIF:\s*(.+)', bloque, re.IGNORECASE)
        solvencia_data['RIF'] = rif_match.group(1).strip() if rif_match else None
        
        # 4. Dirección Fiscal
        dir_match = re.search(r'DIRECCIÓN FISCAL:?\s*(.+)', bloque, re.IGNORECASE)
        solvencia_data['DIRECCION_FISCAL'] = dir_match.group(1).strip() if dir_match else None
        
        # 5. Número de Patente asociado
        patente_match = re.search(r'NÚMERO DE PATENTE:?\s*(.+)', bloque, re.IGNORECASE)
        solvencia_data['N_PATENTE'] = patente_match.group(1).strip() if patente_match else None
        
        # 6. Comprobante de Ingreso
        comp_match = re.search(r'(COMPROBANTE DE INGRESO [^\n]+)', bloque, re.IGNORECASE)
        solvencia_data['COMPROBANTE_INGRESO'] = comp_match.group(1).strip() if comp_match else "N/A"
        
        # 7. Concepto de Pago
        pago_match = re.search(r'(PAGO POR [^\n]+)', bloque, re.IGNORECASE)
        solvencia_data['CONCEPTO_PAGO'] = pago_match.group(1).strip() if pago_match else None
        
        # 8. Fecha de Emisión
        fecha_match = re.search(r'PUERTO CUMAREBO[;\,]\s*([^\n]+)', bloque, re.IGNORECASE)
        if fecha_match:
            raw_date = fecha_match.group(1).strip(' .')
            solvencia_data['FECHA_TEXTO'] = "PUERTO CUMAREBO, " + raw_date
            
            parts_match = re.search(r'(\d{1,2})\s+DE\s+([A-Z]+)\s+DEL?\s+(\d{4})', raw_date, re.IGNORECASE)
            if parts_match:
                dia = parts_match.group(1).zfill(2)
                mes_str = parts_match.group(2).upper()
                anio = parts_match.group(3)
                solvencia_data['FECHA_ISO'] = f"{anio}-{meses.get(mes_str, '01')}-{dia}"
            else:
                solvencia_data['FECHA_ISO'] = None
        else:
            solvencia_data['FECHA_TEXTO'] = None
            solvencia_data['FECHA_ISO'] = None
            
        # 9. Período de Validez
        validez_match = re.search(r'VALIDA HASTA EL\s*(.+)', bloque, re.IGNORECASE)
        solvencia_data['VALIDO_HASTA'] = validez_match.group(1).strip() if validez_match else None
        
        lista_solvencias.append(solvencia_data)
        
    return lista_solvencias

def guardar_en_csv(data, filename="solvencias_extraidas.csv"):
    if not data:
        print("No se encontraron datos para guardar.")
        return
    columnas = data[0].keys()
    with open(filename, mode='w', encoding='utf-8-sig', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=columnas)
        writer.writeheader()
        writer.writerows(data)
    print(f"¡Datos guardados exitosamente en '{filename}'!")

if __name__ == "__main__":
    try:
        with open('solvencias.md', 'r', encoding='utf-8') as f:
            contenido_markdown = f.read()
        datos_procesados = extraer_solvencias(contenido_markdown)
        print("Muestra del primer registro extraído:")
        print(json.dumps(datos_procesados[0], indent=4, ensure_ascii=False))
        print("-" * 50)
        guardar_en_csv(datos_procesados)
    except FileNotFoundError:
        print("Por favor, guarda el texto de origen en un archivo llamado 'solvencias.md'.")