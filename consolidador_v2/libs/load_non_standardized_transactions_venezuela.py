from datetime import datetime
from project_types import Transaction
from openpyxl import load_workbook

def load_non_standardized_transactions_venezuela(path: str, only_withdrawals: bool = False) -> list[Transaction]:
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

  file_name = path

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

    standardized_payment = Transaction(
      date=date,
      reference=str(row[1]).strip(),
      description=row[2].strip(),
      amount=amount,
      bank="Banco de Venezuela",
      account_number="1892"
    )

    if only_withdrawals:
      if standardized_payment.amount < 0:
        payments_list.append(standardized_payment)

      continue


    is_biopago_deposit = "LIQUIDACION TDD BIOPAGOBDV".lower() in standardized_payment.description.lower().strip()

    if standardized_payment.amount > 0 and not is_biopago_deposit:
      payments_list.append(standardized_payment)

  return payments_list

if __name__ == "__main__":
  import os
  import sys
  import json

  path = 'datos/account_statements/25-08-1892.xlsx'

  if not os.path.exists(path):
    print(f"Warning: file {path} doesn't exist")
    sys.exit(1)

  payments_list = load_non_standardized_transactions_venezuela(path)
  print(json.dumps([p.model_dump() for p in payments_list], indent=2, default=str))