from datetime import datetime, timedelta
import json
import sys 
import os 
import openpyxl


from load_non_standardized_transactions_venezuela import load_non_standardized_transactions_venezuela
from load_standardized_transactions import load_standardized_transactions
from load_account_statement_data import load_account_statement_data
from project_types import Transaction

from settings import *

def build_conciliation_report(transactions: list[Transaction], month: int, year: int, bank_account: str, final_amount: float) -> None:
  '''
    This function will build the consolidation report

    Args:
      month
      year
      bank_account

      transactions
      
    Returns:
      None
  '''
  pass

  # pick the template from ./templates/conciliation_report.xlsx

  # load the template
  wb = openpyxl.load_workbook(filename="./templates/conciliation_report.xlsx")
  ws = wb.active

  # insert new rows after row 16
  ws.insert_rows(18, len(transactions) -1 )
  print(f"Inserted {len(transactions)} rows after row 16")

  ws.cell(row=6, column=1).value = f"CONCILIACION BANCO DE {BANKS[bank_account]['name']}"
  ws.cell(row=7, column=1).value = f"CUENTA Nº {BANKS[bank_account]['account_number']} (INGRESOS)"

  ws.cell(row=9, column=5).value = f"{MONTHS_IN_SPANISH[month - 1]} {year}"
  ws.cell(row=10, column=1).value = f"BANCO: {BANKS[bank_account]['name']}"
  ws.cell(row=10, column=3).value = f"C.C. Nº {BANKS[bank_account]['account_number']}"
  ws.cell(row=12, column=4).value = final_amount

  ws.cell(row=11, column=3).value = f"{BANKS[bank_account]['account_number']}"
  ws.cell(row=12, column=3).value = f"{BANKS[bank_account]['account_number']}"

  last_day_of_month = datetime(year, month+1, 1) - timedelta(days=1)
  ws.cell(row=10, column=6).value = last_day_of_month.strftime("%d/%m/%Y")

  ws.cell(row=10, column=6).value = last_day_of_month.strftime("%d/%m/%Y")

  # from A17
  # for each transaction, write the data in the corresponding columns
  for index, transaction in enumerate(transactions, start=16):
    ws.cell(row=index+1, column=1, value=transaction.date.strftime("%d/%m/%Y"))
    ws.cell(row=index+1, column=2, value=transaction.reference[-6:])
    ws.cell(row=index+1, column=3, value=transaction.description)
    ws.cell(row=index+1, column=7, value=transaction.amount)

    # for each cell, copy the format from row 17
    for col in range(1, 8):
      source_cell = ws.cell(row=17, column=col)
      target_cell = ws.cell(row=index+1, column=col)
      if source_cell.has_style:
        target_cell._style = source_cell._style

  # write the sum of the amounts in cell I17 + len(transactions) + 1
  sum_cell_row = 17 + len(transactions)
  sum_cell_column = 7
  sum_cell = ws.cell(row=sum_cell_row, column=sum_cell_column)
  sum_cell.value = f"=SUM(G17:G{17 + len(transactions) - 1})"

  export_folder = "./datos/exports/conciliation_reports/"
  if not os.path.exists(export_folder):
    os.makedirs(export_folder)

  file_name = f"./datos/exports/conciliation_reports/conciliation_report_{bank_account}_{month}_{year}.xlsx"
  wb.save(file_name)
  print(f"File {file_name} generated with the incomes book")
  

if __name__ == "__main__":
  if len(sys.argv) != 5:
    raise ValueError("Must provide month (1-12), year (ex: 2024, 2025), bank account as arguments (1892 or 9290), and relative path to payments file (ex: ./datos/exports/payments.xlsx)")

  month = int(sys.argv[1])
  year = int(sys.argv[2])
  bank_account = sys.argv[3]
  path_to_payments = sys.argv[4]

  account_statement_data = load_account_statement_data(month, year, bank_account)

  transactions = load_standardized_transactions("./datos/exports/payments.xlsx")

  pending_payments = []
  settled_payments = []

  for p in transactions:
    date = p.date 
    settlement_date = p.settlement_date

    if p.account_number != bank_account:
      continue

    if p.matched_settlement_code is not None and p.settlement_date.month == month and p.settlement_date.year == year:
      settled_payments.append(p)
      continue

    if date.month == month and date.year == year:
      pending_payments.append(p)

    

  settled_payments.sort(key=lambda p: p.matched_settlement_code)
  pending_payments.sort(key=lambda p: p.date)

  # print("Pending Payments:")
  # for p in pending_payments:
  #   for p in pending_payments:
  #     print(json.dumps(p.model_dump(), indent=2, default=str))

  print(f'Total Pending Payments: {len(pending_payments)}')

  transactions.sort(key=lambda p: p.settlement_date if p.settlement_date else p.date)

  account_statement_data = load_account_statement_data(month, year, bank_account)

  build_conciliation_report(pending_payments, month, year, bank_account, account_statement_data.final_amount)
