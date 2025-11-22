from openpyxl import load_workbook
from datetime import datetime

from project_types import Transaction

def load_standardized_payments(path: str) -> list[Transaction]:
  """
    If the file doesn't exists, print a warning and return an empty list

    Parameters
    ----------
    path : str
        The path to the standardized payments file

    Returns
    -------
    list
        A list of payments that has an standardized format
  """

  try:
    workbook = load_workbook(path)
  except FileNotFoundError:
    print(f"Warning: file {path} not found")
    return []

  sheet = workbook["Sheet"]
  payments_list = []

  for index, row in enumerate(sheet.iter_rows(values_only=True), start=1):

    if index == 1:
      continue

    date = row[0]
    if type(date) is str:
      try:
        date = datetime.strptime(row[0], "%d/%m/%Y").date()
      except ValueError:
        print(f"Warning: invalid date format {row[0]} in file {path}, row {index}")
        continue
    else:
      date = row[0].date()

    settlement_date = None
    if row[7] is not None:
      if type(row[7]) is str:
        try:
          settlement_date = datetime.strptime(row[7], "%d/%m/%Y").date()
        except ValueError:
          print(f"Warning: invalid settlement date format {row[7]} in file {path}, row {index}")
      elif type(row[7]) is datetime:
        settlement_date = row[7].date()

    payment_obj = Transaction(
      date=date,
      reference=str(row[1]).strip(),
      description=row[2],
      amount=row[3],
      bank=row[4],
      account_number=str(row[5]),

      matched_settlement_code=str(row[6]).strip() if row[6] is not None else None,
      settlement_date=settlement_date,
    )

    # this should not happend, standarized payments should have all the fields, but still...
    if len(row) > 8:
      payment_obj.settlement_description = row[8]

    payments_list.append(payment_obj)

  return payments_list

if __name__ == "__main__":
  import os
  import sys
  import json

  path = 'datos/exports/2025-11-payments.xlsx'

  if not os.path.exists(path):
    print(f"Warning: file {path} doesn't exist")
    sys.exit(1)

  payments = load_standardized_payments(path)
  print(json.dumps([p.model_dump() for p in payments], indent=2, default=str))