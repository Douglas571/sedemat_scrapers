from openpyxl import load_workbook
from datetime import datetime

from project_types import Payment

def load_standarized_payments(path: str) -> list[Payment]:
  """
    If the file doesn't exists, print a warning and return an empty list

    Parameters
    ----------
    path : str
        The path to the standarized payments file

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

    payment_obj = Payment(
      date=row[0].strftime("%Y-%m-%d"),
      reference=str(row[1]).strip(),
      description=row[2],
      amount=row[3],
      bank=row[4],
      account_number=str(row[5]),

      matched_settlement_code=str(row[6]).strip() if row[6] is not None else None,
      settlement_date=row[7].strftime("%Y-%m-%d") if row[7] is not None else None,
    )

    # this should not happend, standarized payments should have all the fields, but still...
    if len(row) > 8:
      payment_obj.settlement_description = row[8]

  return payments_list