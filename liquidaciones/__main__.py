import openpyxl
import csv
import re
from datetime import date



def read_excel_file(file):
    """
    This function read an excel file and return a list of dictionaries with the settlements information
    """
    libro = openpyxl.load_workbook(file, data_only=True)
    liquidaciones = []
    for hoja in libro.worksheets[1:]:
        comprobante = {}
        texto_comprobante = hoja['B8'].value
        comprobante['razon_social'] = hoja['C12'].value
        comprobante['rif'] = f"{hoja['H12'].value or ''} {hoja['H13'].value or ''}".strip()

        cell_value = hoja['B8'].value
        if cell_value:
            match = re.search(r"COMPROBANTE DE INGRESO N°(\d+)", str(cell_value), re.IGNORECASE)
            if match:
                comprobante['num_comprobante'] = match.group(1)

        comprobante['partida'] = extraer_partida(hoja['A21'].value)
        texto_fecha_pago = hoja['C39'].value
        comprobante['fecha_pago'] = texto_fecha_pago
        texto_fecha_liquidacion = hoja['B10'].value
        comprobante['fecha_liquidacion'] = extraer_fecha(texto_fecha_liquidacion)
        comprobante['codigo_banco'] = str(hoja['C37'].value).strip()[-4:]
        comprobante['banco'] = hoja['C36'].value.split(' ')[-1].strip() if hoja['C36'].value is not None else None
        comprobante['referencia'] = hoja['C40'].value
        comprobante['monto'] = hoja['B17'].value
        liquidaciones.append(comprobante)
    return liquidaciones

def extraer_partida(texto):
    """
    This function takes a string in the format "301020700 - PATENTE DE INDUSTRIA Y COMERCIO" and returns the second part
    """
    partes = texto.split('-')
    return partes[1].strip()

def extraer_fecha(texto):
    """
    This function takes a string in the format "dd/mm/yyyy" and returns a datetime object
    """
    
    match = re.search(r"(\d{1,2}) DE (\w+) (\d{4})", texto, re.DOTALL)
    if match:
        dia = int(match.group(1))
        mes_nombre = match.group(2)
        anio = int(match.group(3))
        meses = {
            'ENERO': 1,
            'FEBRERO': 2,
            'MARZO': 3,
            'ABRIL': 4,
            'MAYO': 5,
            'JUNIO': 6,
            'JULIO': 7,
            'AGOSTO': 8,
            'SEPTIEMBRE': 9,
            'OCTUBRE': 10,
            'NOVIEMBRE': 11,
            'DICIEMBRE': 12
        }
        mes = meses[mes_nombre]
        return date(anio, mes, dia)
    else:
        return None

def write_csv_file(liquidaciones, file):
    """
    This function takes a list of dictionaries and writes them to a csv file
    """
    with open(file, 'w', newline='', encoding="utf-8") as csvfile:
        fieldnames = ['num_comprobante', 'razon_social', 'rif', 'partida', 'fecha_pago', 'fecha_liquidacion', 'codigo_banco', 'banco', 'referencia', 'monto']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(liquidaciones)

def write_excel_file(liquidaciones, file):
    """
    This function takes a list of dictionaries and writes them to an excel file
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Liquidaciones"
    fieldnames = ['RAZON SOCIAL', 'CÉDULA O RIF', 'NRO COMPROBANTE', 'PARTIDA', 'FECHA DE PAGO', 'FECHA DE LIQUIDACIÓN', 'COD. BANCO', 'BANCO', 'REFERENCIA', 'MONTO']
    ws.append(fieldnames)
    for liquidacion in liquidaciones:
        ws.append(list(liquidacion.values()))
    wb.save(file)


import sys

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <input_file> <output_file>")
        sys.exit(1)

    archivo_entrada = sys.argv[1]
    archivo_salida_excel = sys.argv[2]
    liquidaciones = read_excel_file(archivo_entrada)

    # archivo_salida = archivo_salida_excel.replace('.xlsx', '-cuadro.csv')
    # write_csv_file(liquidaciones, archivo_salida)

    write_excel_file(liquidaciones, archivo_salida_excel)

