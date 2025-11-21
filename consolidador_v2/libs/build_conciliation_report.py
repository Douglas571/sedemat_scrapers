import json
import sys 
import os 
import openpyxl


from load_non_standardized_payments_venezuela import load_non_standardized_payments_venezuela
from load_standardized_payments import load_standardized_payments
from load_account_statement_data import load_account_statement_data
from project_types import Payment

def build_conciliation_report(transactions: list[Payment], month: int, year: int, bank_account: str) -> None:
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

  # pick the template from ./templates/consolidation_report.xlsx

if __name__ == "__main__":
  if len(sys.argv) != 5:
    raise ValueError("Must provide month (1-12), year (ex: 2024, 2025), bank account as arguments (1892 or 9290), and relative path to payments file (ex: ./datos/exports/payments.xlsx)")

  month = int(sys.argv[1])
  year = int(sys.argv[2])
  bank_account = sys.argv[3]
  path_to_payments = sys.argv[4]

  account_statement_data = load_account_statement_data(month, year, bank_account)

  transactions = load_standardized_payments("./datos/exports/payments.xlsx")

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

  print("Pending Payments:")
  for p in pending_payments:
    for p in pending_payments:
      print(json.dumps(p.model_dump(), indent=2, default=str))

  print(f'Total Pending Payments: {len(pending_payments)}')

  transactions.sort(key=lambda p: p.settlement_date if p.settlement_date else p.date)

  # build_incomes_book(transactions, month, year, bank_account, initial_amount=account_statement_data.initial_amount)
