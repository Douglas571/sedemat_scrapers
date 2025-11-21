from load_standardized_payments import *
from load_non_standardized_payments_venezuela import *

import os

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

from project_types import Payment
from datetime import datetime

MONTHS_IN_SPANISH = [
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SEPTIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE"
]

BANKS = {
  '1892': {
    'name': 'VENEZUELA',
    'account_number': '0102-0339-2500-0107-1892'
  }
}

def build_incomes_book(transactions: list[Payment], month: int, year: int,bank_account: str, initial_amount: float) -> None:
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
    ws.cell(row=row_num, column=1).value = payment.date.strftime("%d/%m/%Y")
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

  

  # if bank account is 1892, pick the ./templates/book_1892.xlsx
  # if it's 9290, pick the ./templates/book_9290.xlsx
  # if not, throw an error and exit  

  # get the payments paid in the given month and year, for this, filter the payments which settlement_date is the given month

  # define a list of dictionaries with the payments data needed for the book
  # for the payments, the list 

  # calculate the 

if __name__ == "__main__":

  month = 4
  year = 2025
  bank_account = '1892'
  transactions = load_standardized_payments("./datos/exports/payments.xlsx")

  settled_payments = [p for p in transactions if p.matched_settlement_code is not None and p.settlement_date.month == month and p.settlement_date.year == year]

  settled_payments.sort(key=lambda p: p.matched_settlement_code)

  commissions = load_non_standardized_payments_venezuela(f"./datos/account_statements/{str(year)[-2:]}-{month:02d}-{bank_account}.xlsx", only_withdrawals=True)

  transactions = settled_payments + commissions

  transactions.sort(key=lambda p: p.settlement_date if p.settlement_date else p.date)

  build_incomes_book(transactions, month, year, bank_account, initial_amount=192482.5)
