import os
import sys

from libs.load_standardized_transactions import *
from libs.load_non_standardized_transactions_venezuela import *
from load_account_statement_data import load_account_statement_data

# stage 7 - Presentation
# this script will build the incomes book

"""

args: 
  month: the month to work with
  year: the year to work with, must have 2 digits
  bank_account: 1892 or 9290

  input_payments: path to the payments file
  input_settlements: path to the settlements file

"""

import openpyxl

from project_types import Transaction
from datetime import datetime

from settings import *

def build_incomes_book(transactions: list[Transaction], month: int, year: int,bank_account: str, initial_amount: float) -> None:
  """
    This function will build the incomes book

    Args:
      transactions: a list of payments

      month: the month to work with, must be between 1 and 12
      year: the year to work with, must have 2 digits
      bank_account: 1892 or 9290

      initial_amount: the initial amount for the account 

    Returns:
      None
  """
  
  # template_path = f"./templates/incomes_book_{bank_account}.xlsx"
  template_path = f"./templates/incomes_book.xlsx"
  wb = openpyxl.load_workbook(template_path)
  ws = wb['incomes_book']


  row_num = 12
  for payment in transactions:

    if payment.account_number != str(bank_account):
      continue

    date = payment.settlement_date if payment.settlement_date else payment.date

    ws.cell(row=row_num, column=1).value = date.strftime("%d/%m/%Y")
    ws.cell(row=row_num, column=2).value = payment.description
    ws.cell(row=row_num, column=3).value = payment.reference[-6:]
    ws.cell(row=row_num, column=4).value = payment.matched_settlement_code
    if payment.amount > 0:
      ws.cell(row=row_num, column=5).value = payment.amount
    else:
      ws.cell(row=row_num, column=6).value = payment.amount
    row_num += 1

  ws.cell(row=11, column=7).value = initial_amount
  ws.cell(row=11, column=1).value = MONTHS_IN_SPANISH[month - 1]
  ws.cell(row=11, column=2).value = f"SALDO INICIAL AL 01/{month:02d}/{year}"

  ws.cell(row=6, column=1).value = f"LIBRO DEL BANCO {BANKS[bank_account]['name']}"
  ws.cell(row=7, column=1).value = f"CUENTA Nº {BANKS[bank_account]['account_number']} (INGRESOS)"

  export_folder = "./datos/exports/incomes_books"
  if not os.path.exists(export_folder):
    os.makedirs(export_folder)

  file_name = f"./datos/exports/incomes_books/incomes_book_{bank_account}_{month}_{year}.xlsx"
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

  settled_payments = [p for p in transactions if p.matched_settlement_code is not None and p.settlement_date.month == month and p.settlement_date.year == year]

  settled_payments.sort(key=lambda p: p.matched_settlement_code)

  commissions = load_non_standardized_transactions_venezuela(f"./datos/account_statements/{str(year)[-2:]}-{month:02d}-{bank_account}.xlsx", only_withdrawals=True)

  transactions = settled_payments + commissions

  transactions.sort(key=lambda p: p.settlement_date if p.settlement_date else p.date)

  build_incomes_book(transactions, month, year, bank_account, initial_amount=account_statement_data.initial_amount)
