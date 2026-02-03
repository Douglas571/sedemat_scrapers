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

import re
from libs.settings import SHOW_WARNINGS

import os
from openpyxl import load_workbook
import json
from datetime import datetime, date, timedelta
import sys

def parse_date_intervals(data):
    # Regex Breakdown:
    # (\d{1,2})  -> Capture Group 1: Day (1 or 2 digits)
    # [-\/]      -> Separator: Matches either hyphen or forward slash
    # (\d{1,2})  -> Capture Group 2: Month (1 or 2 digits)
    # [-\/]      -> Separator: Matches either hyphen or forward slash
    # (\d{2,4})  -> Capture Group 3: Year (2 to 4 digits)
    date_pattern = re.compile(r"(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})")
    
    results = []

    matches = date_pattern.findall(data)
        
    parsed_dates = []
    
    for day, month, year_str in matches:
      try:
        d, m = int(day), int(month)
        y = int(year_str)
        
        # Handle 2-digit years (e.g., '25' becomes 2025)
        if y < 100:
            y += 2000
        
        # Create a datetime object for accurate comparison
        dt = datetime(y, m, d)
        parsed_dates.append(dt)
      except ValueError:
        # Skips invalid dates (e.g. if a month is 13)
        continue
    
    # We need at least one date to form a range (min and max can be the same)
    if parsed_dates:
      min_date = min(parsed_dates)
      max_date = max(parsed_dates)

      return [min_date.date(), max_date.date()]
    
    return None

def load_settlements_list(file_path):
  settlements_list = []
  workbook = load_workbook(file_path)
  sheet = workbook.active
  for index, row in enumerate(sheet.iter_rows(values_only=True), start=1):

    if index == 1:
      continue

    settlement = {}
    settlement['legal_name'] = row[0]
    settlement['rif_cedula'] = row[1]
    settlement['num_comprobante'] = str(row[2]).strip()
    settlement['pago_por'] = row[3]
    settlement['fecha_pago'] = row[4]
    settlement['fecha'] = row[5]
    settlement['cuenta'] = row[6]
    settlement['banco'] = row[7]
    
    settlement['referencia'] = str(row[8]).strip()
    settlement['monto'] = row[9]
    settlement['payments'] = []

    settlement['not_found_payments'] = ''

    settlement['amount_difference'] = 0

    settlement['is_verified'] = True
    settlement['is_voided'] = False

    settlement['fecha_pago_rango'] = row[3]

    is_exonerated = any([
      'EXONERADO' in str(x).upper() if x is not None else False
      for x in [settlement['referencia'], settlement['monto'], settlement['banco'], settlement['cuenta']]
    ]) or settlement['referencia'] == "None"

    settlement['is_exonerated'] = is_exonerated

    if not is_exonerated:

      references = settlement['referencia'].replace("/", " ").replace("-", " ").split(" ")
      references = [ref.strip() for ref in references if ref.strip() != '']
      settlement['referencia'] = '-'.join(references)

      paid_at = settlement['fecha_pago']
      if type(paid_at) == str and '/' in paid_at and '-' in paid_at:
        # parse date intervals
        parsed_dates = parse_date_intervals(paid_at)
        
        paid_at = parsed_dates

      elif type(paid_at) == str:
        formats = ["%d/%m/%Y", "%Y-%m-%d %H:%M:%S"]
        for fmt in formats:
          try:
            paid_at = datetime.strptime(paid_at, fmt).date()
            break
          except ValueError:
            pass
            # print(f"Warning: Could not parse paid_at date '{paid_at_raw}' with format {fmt}")
      elif type(paid_at) == datetime:
        paid_at = paid_at.date()
          

      settlement['fecha_pago'] = paid_at

      for ref in references:
        payment = {}
        payment['reference'] = ref
        payment['amount'] = row[9] if len(references) == 1 else None
        payment['not_found'] = True
        settlement['payments'].append(payment)

    if is_exonerated:
      settlement['payments'] = []
      settlement['not_found_payments'] = ''

      
    settlements_list.append(settlement)
  return settlements_list

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

        payment_date = row[0]
        if type(payment_date) is datetime:
          payment_date = payment_date.date()

        payment = {}
        payment["date"] = payment_date
        payment["reference"] = str(row[1]).strip()
        payment["description"] = row[3]
        
        # payment["amount"] = float(row[4] or row[5])
        debit = float(row[4] or 0) * -1
        credit = float(row[5] or 0)
        payment["amount"] = debit or credit

        payment["bank"] = "BDT"
        payment["account_number"] = "9290"

        if payment['amount'] > 0: 
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

  sheet = workbook["data"]
  payments_list = []

  for index, row in enumerate(sheet.iter_rows(values_only=True), start=1):

    if index == 1:
      continue

    # if row[4] is an integer or float, pass it as it is, else, convert 
    amount = 0
    try:
      amount = float(row[4].replace(".", "").replace(",", "."))
    except AttributeError:
      amount = float(row[4] or 0)

    date = row[0]
    if type(date) is str:
      try:
        date = datetime.strptime(row[0], "%d/%m/%Y").date()
      except ValueError:
        print(f"Warning: invalid date format {row[0]} in file {file_name}, row {index}")
        continue
    else:
      date = row[0].date()

    payment = {
      "date": date,
      "reference": str(row[1]).strip(),
      "description": row[2],
      "amount": amount,
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

    # omit empty rows (those with amount = None)
    if row[4] == None: 
      continue

    # print(row)
    payment = {
      "date": row[1].date(),
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


def dates_are_close_by(date1, date2, days):
  """
    Compare two dates to see if they are close by a given number of days

    Args:
      date1 (datetime): the first date
      date2 (datetime): the second date
      days (int): the number of days to compare

    Returns:
      bool: True if the dates are close by the given number of days, False otherwise
  """
  # print(f"date1: {date1}, date2: {date2}")
  return abs((date1 - date2).days) <= days

def is_date_in_range(date, time_range, isBiopago=False):
  """
  Check if a given date is within a given time range

  Args:
    date (datetime): the date to check
    time_range (list): a list with two dates, the start and end of the time range

  Returns:
    bool: True if the date is within the time range, False otherwise
  """
  start_date, end_date = time_range

  if (isBiopago):
    start_date = start_date - timedelta(days=1)
    end_date = end_date + timedelta(days=1)

  return start_date <= date <= end_date

def asigne_payments_to_settlements(payments_list, settlements_list):
  """
    I want to check settlement's payments against the payments_list

    if the settlement payments are not in the payments list, then check the settlement['is_verified'] = false

    if it find the payments, replace it in the settlement['payments'] list with the payment from the payments_list
  """

  counter = 0
  for settlement in settlements_list:

    if settlement['num_comprobante'] == '15':
      print(f"Settlement {settlement['num_comprobante']} found")
      for payment in settlement['payments']:
        print(payment)

    for payment in settlement['payments']:
      found = False

      for payment_in_list in payments_list:

        isBiopago = payment_in_list['bank'] == 'BIOPAGO'

        six_digit_ref = str(payment['reference'])[-6:].zfill(6)
        four_digit_ref = str(payment['reference'])[-4:].zfill(4)

        is_deposit = 'cumarebo' == payment_in_list['description'].lower().strip()

        # the date is valid when payment date is greater or equal to settlement date
        # has_valid_date = payment_in_list['date'] >= settlement['fecha_pago']
        # if not has_valid_date:
        #   print(f"Warning: payment {payment_in_list['reference']} has an invalid date: {payment_in_list['date']} (settlement date: {settlement['fecha']})")

        #   continue

        is_valid_date = False

        if type(settlement['fecha_pago']) == list and len(settlement['fecha_pago']) == 2:
          is_valid_date = is_date_in_range(payment_in_list['date'], settlement['fecha_pago'], isBiopago)
        elif type(settlement['fecha_pago']) == date:
          # print(f"date: {payment_in_list['date']}, reference: {payment_in_list['reference']}, bank: {payment_in_list['bank']}")
          is_valid_date = dates_are_close_by(payment_in_list['date'], settlement['fecha_pago'], 2)


        the_reference_match = (str(payment_in_list['reference']).endswith(six_digit_ref) or (str(payment_in_list['reference']).endswith(four_digit_ref)) and is_deposit) or (settlement['banco'] == 'BDT' and str(payment_in_list['description']).endswith(six_digit_ref))

        if (payment['reference'] == '320585'): 
          # 320585
          # print(f'checking the payment {payment['reference']} in settlment {settlement['num_comprobante']}')
          # print(f"  six_digit_ref: {six_digit_ref}, four_digit_ref: {four_digit_ref}")

          if (payment_in_list['reference'].endswith(six_digit_ref) or payment_in_list['reference'].endswith(four_digit_ref)):
            print(f"PRELIMINAR CHECK: {payment_in_list}")
            print(f"  is_deposit: {is_deposit}")
            print(f'  is_valid_date: {is_valid_date}')
            print(f"    payment_in_list['date']: {payment_in_list['date']}")
            print(f"    settlement['fecha_pago']: {settlement['fecha_pago']}")
            print(f'  the_reference_match: {the_reference_match}')

            print(f"  str(payment_in_list['reference']).endswith(six_digit_ref): {str(payment_in_list['reference']).endswith(six_digit_ref)}")
            print(f"  str(payment_in_list['reference']).endswith(four_digit_ref): {str(payment_in_list['reference']).endswith(four_digit_ref)}")
            print(f"  is_deposit: {is_deposit}")
            print(f"  settlement['banco'] == 'BDT' and str(payment_in_list['description']).endswith(six_digit_ref): {(settlement['banco'] == 'BDT') and (str(payment_in_list['description']).endswith(six_digit_ref))}")

            print('  Settlement info:')
            print(f"    settlement['num_comprobante']: {settlement['num_comprobante']}")

          if (the_reference_match and is_valid_date):
            print(f"found the matching payment_in_list: {payment_in_list}")


        # if ('424505' in payment_in_list['reference'] and settlement['num_comprobante'] == '15'):
        #     print(f"Payment {payment_in_list['reference']} found in settlement {settlement['num_comprobante']}")
        #     print(f"six_digit_ref: {six_digit_ref}, four_digit_ref: {four_digit_ref}")
        #     print(f"is_valid_date: {is_valid_date}")
        #     print(f"payment_in_list['date']: {payment_in_list['date']}")
        #     print(f"settlement['fecha_pago']: {settlement['fecha_pago']}")
        #     print(f"the_reference_match: {the_reference_match}")
        #     print('')

        if the_reference_match and is_valid_date:

          found = True
          payment_index = settlement['payments'].index(payment)
          settlement['payments'][payment_index] = payment | payment_in_list

          settlement['payments'][payment_index]['not_found'] = False

          payment_in_list['matched_settlement'] = settlement['num_comprobante']
          payment_in_list['settlement_date'] = settlement['fecha']
          payment_in_list['settlement_description'] = settlement['legal_name'] + " - " + settlement['rif_cedula'] + " - " + settlement['pago_por']
          counter = counter + 1
          break

        elif the_reference_match and not is_valid_date:
          pass
          # print(f"Warning: payment {payment_in_list['reference']} has a valid reference match with settlement {settlement['num_comprobante']}, but the date is invalid: payment reference: {six_digit_ref or four_digit_ref}, settlement reference: {settlement['referencia']}, payment date: {payment_in_list['date']} (settlement date: {settlement['fecha_pago']}), amount: {payment_in_list['amount']}")
        
      
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

      total_amount = round(total_amount, 2)
      # print(json.dumps(settlement, indent=2, default=str))
      difference = total_amount - (settlement['monto'] or 0)
      if difference != 0 and len(not_found) == 0:
        # print(f"Settlement {settlement['num_comprobante']}: amount: {settlement['monto']}, payment sum: {total_amount}, difference: {difference}")
        settlement['is_verified'] = False
        settlement['amount_difference'] = difference

      # if there is lacking payments, then display the missing amount 
      if (len(not_found) > 0):
        lacking_amount = (settlement['monto'] or 0) - total_amount
        settlement['amount_difference'] = lacking_amount


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

def test_dates_are_close_by():
  print(f"Testing dates_are_close_by:")
  print(f"  assert dates_are_close_by(date(2020, 1, 1), date(2020, 1, 3), 2) => {dates_are_close_by(date(2020, 1, 1), date(2020, 1, 3), 2)}")
  print(f"  assert dates_are_close_by(date(2020, 1, 3), date(2020, 1, 1), 2) => {dates_are_close_by(date(2020, 1, 3), date(2020, 1, 1), 2)}")
  print(f"  assert not dates_are_close_by(date(2020, 1, 1), date(2020, 1, 4), 2) => {not dates_are_close_by(date(2020, 1, 1), date(2020, 1, 4), 2)}")
  print(f"  assert not dates_are_close_by(date(2020, 1, 4), date(2020, 1, 1), 2) => {not dates_are_close_by(date(2020, 1, 4), date(2020, 1, 1), 2)}")
  

def main(argv):

  # SETTLEMENTS_PATH = 'datos/settlements/cuadro_to_use.xlsx'
  SETTLEMENTS_PATH = 'datos/settlements/2026_liquidaciones.xlsx'

  if len(argv) != 3:
    print("Usage: python3 __main__.py <month> <year>\n")
    print("Example: python3 __main__.py 01 2024\n")
    print("Example: python3 __main__.py 02 2025\n")
    print("Example: python3 __main__.py 03 2026\n")
    return

  month = int(argv[1])
  year = int(argv[2])

  payments_list = load_payments_list(month, year)
  settlements_list = load_settlements_list(SETTLEMENTS_PATH)

  # print(json.dumps(payments_list, indent=2, default=str))
  asigne_payments_to_settlements(payments_list, settlements_list)

  # print(json.dumps(settlements_list, indent=2, default=str))
  # print(json.dumps(payments_list, indent=2, default=str))


  if not os.path.exists('datos/exports'):
    os.makedirs('datos/exports')

  export_to_excel(payments_list, f"datos/exports/{year}-{month}-payments.xlsx")
  
  
  
  for settlement in settlements_list:
    if type(settlement['fecha_pago']) == list:
      settlement['fecha_pago'] = '{:%d/%m/%Y} - {:%d/%m/%Y}'.format(* settlement['fecha_pago'])

  export_to_excel(settlements_list, f"datos/exports/{year}-{month}-settlements.xlsx")


if __name__ == "__main__":
  main(sys.argv)
