"""
Settlements Map(
  "1000" => {
    "legal_name": "John Doe",
    "rif_cedula": "123456789",
    "num_comprobante": "1000",
    "pago_por" : "John Doe",
    "fecha_pago": "2020-01-01",
    "fecha": "2020-01-01",
    "cuenta": "1892",
    "banco": "Banco de Venezuela",
    "monto": "10000.00",
    "referencia": "1000 / 1000",
    "payments": [
      {
        "reference": "1000",
        "amount": "10000.00" (or null if the settlement has many payments and the individual amount is not available)
      }
    ]
  },
  ...
)


"""
from openpyxl import load_workbook
import json
from datetime import datetime
import sys

def load_settlements_map():
  settlements_map = {}
  workbook = load_workbook('datos/settlements/cuadro_to_use.xlsx')
  sheet = workbook.active
  for index, row in enumerate(sheet.iter_rows(values_only=True), start=1):

    if index == 1:
      continue

    settlement = {}
    settlement['legal_name'] = row[0]
    settlement['rif_cedula'] = row[1]
    settlement['num_comprobante'] = row[2]
    settlement['pago_por'] = row[3]
    settlement['fecha_pago'] = row[4]
    settlement['fecha'] = row[5]
    settlement['cuenta'] = row[6]
    settlement['banco'] = row[7]
    
    settlement['referencia'] = row[8]
    settlement['monto'] = row[9]
    settlement['payments'] = []

    references = str(settlement['referencia']).strip().replace(" ", "").replace("/", "-").split("-")
    
    for ref in references:
      payment = {}
      payment['reference'] = ref
      payment['amount'] = row[9] if len(references) == 1 else None
      settlement['payments'].append(payment)
    settlements_map[settlement['num_comprobante']] = settlement
  return settlements_map

# print(json.dumps(load_settlements_map(), indent=2, default=str))


def load_9290_payments(month, year):
    """
    look for the back account statements for the given month and year 

    there are 2 files 

    {month}-{year}-9290.xlsx
    {month}-{year}-1892.xlsx

    call the respective function for each file and merge the results 

    return a list of dictionaries 

    Parameters
    ----------
    month : str
      The month of the year
    year : int
      The year

    Returns
    -------
    list
      A list of dictionaries containing the payments for the given month and year
    """
    file_name = f"datos/account_statements/{year}-{month}-9290.xlsx"

    try:
      workbook = load_workbook(file_name)
      sheet = workbook["Table 2"]
      payments_list = []

      for index, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        payment = {}
        payment["date"] = row[0]
        payment["reference"] = row[1]
        payment["description"] = row[2]
        payment["amount"] = float(row[3] or row[4])
        payment["bank"] = "BDT"
        payment["account_number"] = "9290"

        payments_list.append(payment)

      return payments_list
    except FileNotFoundError:
      print(f"Warning: file {file_name} not found")
      return []

def load_1892_payments(month, year):
  """
    If the file doesn't exists, print a warning and return an empty list

    Parameters
    ----------
    month : str
        The month of the year
    year : int
        The year

    Returns
    -------
    list
        A list of payments for the given month and year
  """

  file_name = f"datos/account_statements/{year}-{month}-1892.xlsx"

  try:
    workbook = load_workbook(file_name)
  except FileNotFoundError:
    print(f"Warning: file {file_name} not found")
    return []

  sheet = workbook.active
  payments_list = []

  for index, row in enumerate(sheet.iter_rows(values_only=True), start=1):

    if index == 1:
      continue

    payment = {
      "date": datetime.strptime(row[0], "%d/%m/%Y"),
      "reference": row[1],
      "description": row[2],
      "amount": float(row[3].replace(".", "").replace(",", ".")),
      "bank": "Banco de Venezuela",
      "account_number": "1892"
    }

    payments_list.append(payment)

  return payments_list

def load_biopago_payments(month, year):
  file_name = f"datos/account_statements/{year}-{month}-biopago.xlsx"

  try:
    workbook = load_workbook(file_name)
  except FileNotFoundError:
    print(f"Warning: file {file_name} not found")
    return []

  sheet = workbook.active
  payments_list = []

  for index, row in enumerate(sheet.iter_rows(values_only=True), start=1):

    if index == 1:
      continue

    print(row)
    payment = {
      "date": row[1],
      "reference": row[0],
      "description": f"{row[7]} {row[5]}",
      "amount": float(row[4]),
      "bank": "BIOPAGO",
      "account_number": "1892"
    }
    payments_list.append(payment)

  return payments_list
  

def load_payments_list(month, year):
  """
    look for the back account statements for the given month and year 

    there are 3 files 

    {month}-{year}-9290.xlsx
    {month}-{year}-1892.xlsx
    {month}-{year}-biopago.xlsx

    call the respective function for each file and merge the results 

    return a list of dictionaries 
  """

  month = datetime.now().month if not month else month
  month = "0" + str(month) if len(str(month)) == 1 else str(month)

  current_year = datetime.now().year
  if not year:
    year = current_year
  elif not (2024 <= year <= current_year):
    raise ValueError(f"Year must be between 2024 and {current_year}")
  
  year = str(year)[-2:]
  
  print(f"month: {month}, year: {year}")

  payments_list = []
  payments_list.extend(load_9290_payments(month, year))
  payments_list.extend(load_1892_payments(month, year))
  payments_list.extend(load_biopago_payments(month, year))
  return payments_list

def main(argv):
  if len(argv) != 3:
    print("Usage: python3 consolidador_v2.py <month> <year>")
    return

  month = int(argv[1])
  year = int(argv[2])

  print(json.dumps(load_payments_list(month, year), indent=2, default=str))

if __name__ == "__main__":
  main(sys.argv)
