import locale
import pandas as pd
import os
from datetime import datetime

locale.setlocale(locale.LC_ALL, 'es_VE.utf8')

# this program will check taht that the payments in account statement files, are in the settlement files for that month 
# 
# get for console input 
#   month (it can be a number from 1 to 12) 
#   year (it can be a number from 2020 to current year)

# ask user for month and year
month = input("Enter month (1-12): ")
year = input("Enter year (2020-current): ")

# validate month and year
current_year = datetime.now().year
current_month = datetime.now().month

if not year:
    year = current_year
else:
    year = int(year)

if not month:
    month = current_month
else:
    month = int(month)


if not (1 <= month <= 12):
    raise ValueError("Month must be between 1 and 12")

if not (2020 <= year <= current_year):
    raise ValueError(f"Year must be between 2020 and {current_year}")

# format month and year for filenames
mm = f"{month:02d}"
yy = str(year)[-2:]

# --------------------------
#          LOAD FASE
# --------------------------

# after you get the month and year, load the account statements for the accounts 1892 and 9290
account_statements_dir = "./datos/account_statements"

file_9290 = os.path.join(account_statements_dir, f"{mm}-{yy}-9290.xlsx")
file_1892 = os.path.join(account_statements_dir, f"{mm}-{yy}-1892.xlsx")

# for 9290
#  look into the sheet "Table 2"
#  it has the following columns
#  1. Fecha
#  2. Referencia
#  3. Código 
#  4. Descripción
#  5. Débito
#  6. Crédito
#  7. Saldo
df_9290 = pd.read_excel(file_9290, sheet_name="Table 2")

# for 1892
#  look into the sheet "data"
#  it has the following columns
#  1. fecha
#  2. referencia
#  3. concepto 
#  4. saldo
#  5. month
#  6. tipoMovimiento
#  7. rif
#  8. numeroCuenta
df_1892 = pd.read_excel(file_1892, sheet_name="data")

# for biopago transactions, load the file from ./datos/account_statements/{MM}-{YY}-biopago.xlsx
# it has the following columns
# Nro.	Fecha	Instrumento	Emisor	Monto	Equipo	Lote	Cédula Pagador	Resultado	Autorización
biopago_file = f"./datos/account_statements/{mm}-{yy}-biopago.xlsx"
df_biopago = pd.read_excel(biopago_file, header=None, names=[
    "number",
    "date",
    "instrument",
    "issuer",
    "amount",
    "equipment",
    "lot",
    "payer_id",
    "result",
    "authorization"
])

print("Biopago transactions loaded:", df_biopago.shape)
print(df_biopago.to_string())

# load the settlments 
# the file is in ./datos/settlements/cuadro-{MM}-{YY}.xlsx
# it has the following columns
# 1. razon_social
# 2. rif_cedula
# 3. num_comprobante
# 4. pago_por
# 5. fecha_pago
# 6. fecha
# 7. cuenta
# 8. banco
# 9. referncia
# 10. monto
# settlements_file = f"./datos/settlements/cuadro-{mm}-{yy}.xlsx"
settlements_file = f"./datos/settlements/cuadro-{mm}-{yy}.xlsx"
df_settlements = pd.read_excel(settlements_file)

# in this case, remove all the settlements that has "EXONERADO" in refernece column
# df_settlements = df_settlements[~df_settlements['referncia'].astype(str).str.contains("EXONERADO", case=False, na=False)]

# now you have:
# df_9290 → account 9290 statement
# df_1892 → account 1892 statement
# df_settlements → settlements (filtered)

# here you would implement the logic to check that payments in account statements are in settlements
# (matching by reference, amount, or other criteria depending on your business rules)
print("Account 9290 statement loaded:", df_9290.shape)
print("Account 1892 statement loaded:", df_1892.shape)
print("Settlements loaded (filtered):", df_settlements.shape)

# print("Data in account 1892 statement:")
# print(df_1892.to_string())

# print("Data in account 9290 statement:")
# print(df_9290.to_string())

# ------------------------------------------
#             NORMALIZATION FASE 
# ------------------------------------------

# the common structure for data will be
#  1. reference 
#  2. amount
#  3. date
#  4. bank
#  5. account_number 

# map the data in df_9290 and df_1892 into the common structure
# 9290 => BDV
# 1892 => BANCO DE VENEZUELA

# normalize df_9290
df_9290_norm = pd.DataFrame({
    "reference": df_9290["Referencia"],
    "amount": df_9290["Crédito"].fillna(0) - df_9290["Débito"].fillna(0),  # positive = credit, negative = debit
    "date": pd.to_datetime(df_9290["Fecha"], format="%d-%m-%Y", errors="coerce"),
    "bank": "BDV",
    "account_number": "9290",
    "description": df_9290["Descripción"],
    "settlementCode": '',
    "settlementDate": None
})

# normalize df_1892
df_1892_norm = pd.DataFrame({
    "reference": df_1892["referencia"],
    "amount": df_1892["monto"],  # assuming saldo is the transaction amount
    "date": pd.to_datetime(df_1892["fecha"], format="%d/%m/%Y", errors="coerce"),
    "bank": "BANCO DE VENEZUELA",
    "account_number": "1892",
    "description": df_1892["concepto"],
    "settlementCode": '',
    "settlementDate": None
})

# merge then in a single object
payments = pd.concat([df_9290_norm, df_1892_norm], ignore_index=True)

# print the list of payments
# print("Normalized Payments:")
# print(payments.to_string())


# ------------------------------------------
#             CONSOLIDATION FASE 
# ------------------------------------------


# for each entry in settlements data frame, print reference

not_settled_payments = []

paymentsDict = payments.to_dict(orient="records")

# for each payment
for index, payment in payments.iterrows():
  # for each settlement
  found = False

  payment_reference = str(payment["reference"]).split(".")[0][-6:]

  if isinstance(payment["amount"], str):
    paymentsDict[index]["amount"] = locale.atof(payment["amount"])

  for index_settlement, settlement in df_settlements.iterrows():
    # if payment reference is included in settlement reference, continue with another payment
    if payment_reference in str(settlement["referencia"]):
      found = True
      paymentsDict[index]["settlementCode"] = str(settlement["num_comprobante"])
      paymentsDict[index]["settlementDate"] = str(settlement["fecha"])
  
  # if not, add payment to not settled payments
  if not found:
    not_settled_payments.append(payment)

print("Payments not settled:", len(not_settled_payments))

toPrintData = pd.DataFrame(paymentsDict)
print(toPrintData.to_string())

toPrintData.to_excel(f"./datos/payments_{mm}_{yy}.xlsx", index=False)
print(f"File ./datos/payments_{mm}_{yy}.xlsx generated with {len(payments)} payments")


# generate an excel file with the payments not settled
# not_settled_payments_df = pd.DataFrame(not_settled_payments)
# not_settled_payments_file = f"./datos/not_settled_payments_{mm}_{yy}.xlsx"
# not_settled_payments_df.to_excel(not_settled_payments_file, index=False)
# print(f"File {not_settled_payments_file} generated with {len(not_settled_payments)} payments not settled")
