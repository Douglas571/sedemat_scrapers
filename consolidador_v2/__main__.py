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
import os
from openpyxl import load_workbook
import json
from datetime import datetime
import sys

def load_settlements_list():
  settlements_list = []
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
    
    settlement['referencia'] = str(row[8]).strip()
    settlement['monto'] = row[9]
    settlement['payments'] = []

    settlement['not_found_payments'] = ''



    settlement['is_verified'] = True
    # settlement['comment'] = ''

    references = str(settlement['referencia']).strip().replace(" ", "").replace("/", "-").split("-")

    settlement['is_exonerated'] = 'EXONERADO' in settlement['referencia'].upper() or 'EXONERADO' in str(settlement['monto']).upper() or 'EXONERADO' in settlement['banco'].upper() or 'EXONERADO' in str(settlement['cuenta']).upper()

    if not settlement['is_exonerated']:
      for ref in references:
        payment = {}
        payment['reference'] = ref
        payment['amount'] = row[9] if len(references) == 1 else None
        payment['not_found'] = True
        settlement['payments'].append(payment)
      
    settlements_list.append(settlement)
  return settlements_list

# print(json.dumps(load_settlements_list(), indent=2, default=str))


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

        if index == 1:
          continue

        payment = {}
        payment["date"] = row[0]
        payment["reference"] = str(row[1]).strip()
        payment["description"] = row[3]
        
        # payment["amount"] = float(row[4] or row[5])
        debit = float(row[4] or 0) * -1
        credit = float(row[5] or 0)
        payment["amount"] = debit or credit

        payment["bank"] = "BDT"
        payment["account_number"] = "9290"

        if debit > 0: 
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
      "reference": str(row[1]).strip(),
      "description": row[2],
      "amount": float(row[4].replace(".", "").replace(",", ".")),
      "bank": "Banco de Venezuela",
      "account_number": "1892"
    }

    is_biopago_deposit = "LIQUIDACION TDD BIOPAGOBDV".lower() in payment["description"].lower().strip()

    if payment["amount"] > 0 and not is_biopago_deposit:
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

    # print(row)
    payment = {
      "date": row[1],
      "reference": str(row[0]).strip(),
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

  current_year = datetime.now().year
  if not year:
    year = current_year
  elif not (2024 <= year <= current_year):
    raise ValueError(f"Year must be between 2024 and {current_year}")
  
  year = str(year)[-2:]
  
  print(f"month: {month}, year: {year}")

  payments_list = []
  for given_month  in range(1, month+1):

    string_month = "0" + str(given_month) if len(str(given_month)) == 1 else str(given_month)

    print(f"Loading payments for month: {given_month}, year: {year}")

    payments_list.extend(load_9290_payments(string_month, year))
    payments_list.extend(load_1892_payments(string_month, year))
    payments_list.extend(load_biopago_payments(string_month, year))

  return payments_list

def asigne_payments_to_settlements(payments_list, settlements_list):
  """
    I want to check settlement's payments against the payments_list

    if the settlement payments are not in the payments list, then check the settlement['is_verified'] = false

    if it find the payments, replace it in the settlement['payments'] list with the payment from the payments_list
  """

  counter = 0
  for settlement in settlements_list:
    for payment in settlement['payments']:
      found = False
      for payment_in_list in payments_list:

        six_digit_ref = str(payment['reference'])[-6:].zfill(6)
        four_digit_ref = str(payment['reference'])[-4:].zfill(4)

        is_deposit = 'cumarebo' == payment_in_list['description'].lower().strip()

        if str(payment_in_list['reference']).endswith(six_digit_ref) or (is_deposit and str(payment_in_list['reference']).endswith(four_digit_ref)):
          found = True
          payment_index = settlement['payments'].index(payment)
          settlement['payments'][payment_index] = payment | payment_in_list

          settlement['payments'][payment_index]['not_found'] = False

          payment_in_list['matched_settlement'] = settlement['num_comprobante']
          payment_in_list['settlement_date'] = settlement['fecha']
          counter = counter + 1
          break
      
    if settlement['is_exonerated']:
      settlement['is_verified'] = True

    else:

      not_found = []
      for p in settlement['payments']:
        if p.get('not_found', True):
          not_found.append(str(p.get('reference')))

      settlement['not_found_payments'] = ", ".join(not_found) if not_found else ""
      if not_found:
        settlement['is_verified'] = False
      

      total_amount = sum([payment['amount'] for payment in settlement['payments'] if not payment.get('not_found', True)])

      print(total_amount, settlement['monto'])
      print(json.dumps(settlement, indent=2, default=str))
      
      settlement['is_verified'] = total_amount - settlement['monto'] == 0

  print(f"Payments found: {counter}")


def export_to_excel(list_of_dicts, file_name):
  from openpyxl import Workbook

  workbook = Workbook()
  sheet = workbook.active

  if not list_of_dicts:
    print("No data to export")
    return

  headers = list(list_of_dicts[0].keys())
  sheet.append(headers)

  for item in list_of_dicts:
    item['payments'] = ''
    row = [item.get(header, "") for header in headers]
    sheet.append(row)

  workbook.save(file_name)
  print(f"Data exported to {file_name}")

def main(argv):
  if len(argv) != 3:
    print("Usage: python3 consolidador_v2.py <month> <year>")
    return

  month = int(argv[1])
  year = int(argv[2])

  payments_list = load_payments_list(month, year)
  settlements_list = load_settlements_list()

  # print(json.dumps(payments_list, indent=2, default=str))
  asigne_payments_to_settlements(payments_list, settlements_list)

  print(json.dumps(settlements_list, indent=2, default=str))
  # print(json.dumps(payments_list, indent=2, default=str))


  if not os.path.exists('datos/exports'):
    os.makedirs('datos/exports')

  export_to_excel(payments_list, f"datos/exports/{year}-{month}-payments.xlsx")
  export_to_excel(settlements_list, f"datos/exports/{year}-{month}-settlements.xlsx")


if __name__ == "__main__":
  main(sys.argv)
