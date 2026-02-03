# COMMANDS

## to generate the payments and settlements files
python . <month, ex: 01, 02...> <year, ex: 2025, 2026...>

## to generate consolidation file 
python ./libs/build_consolidation_report.py <month> <year> <bank_account> <path_to_payments>

month: should  be 01, 02, 03...
year: should be 2025, 2026...
bank_account: 9290 or 1892
path_to_payments: ./data/exports/payments_file.xlsx

## to 