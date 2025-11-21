

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

def load_venezuela_commissions_and_withdrawals(month, year):
  """
    this function retrieve a list of commissions and withdrawals from venezuela in a given month with a standard format
  """

def build_book(month, year, bank_account):
  # pick the template to fill 

  # if bank account is 1892, pick the ./templates/book_1892.xlsx
  # if it's 9290, pick the ./templates/book_9290.xlsx
  # if not, throw an error and exit  

  # get the payments paid in the given month and year, for this, filter the payments which settlement_date is the given month

  # define a list of dictionaries with the payments data needed for the book
  # for the payments, the list 

  # calculate the 
