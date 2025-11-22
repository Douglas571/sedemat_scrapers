import openpyxl
from project_types import Settlement


def save_standardized_settlements(settlements_list: list[Settlement], file_name: str) -> None:
    """
    This function will save the Settlement objects into an excel file

    Args:
      settlements_list: a list of Settlement objects
      file_name: the name of the excel file to save the data into

    Returns:
      None
    """

    # load the template
    wb = openpyxl.Workbook()
    ws = wb.active

    # dump into a dict, take the keys as header and append the values into the row
    settlement_dict = {}
    for index, settlement in enumerate(settlements_list):
        settlement_dict[index] = settlement.model_dump()
    
    # write the headers
    headers = list(settlement_dict[0].keys())
    for index, header in enumerate(headers, start=1):
        print(f"{index}: {header}")
    ws.append(headers)

    # write the values
    row_index = 2
    for index, settlement in settlement_dict.items():
        for index, value in enumerate(settlement.values()):
            
          if index == 14: # Omit the payments column
              continue

          # print(row_index, index+1, value)
          ws.cell(row_index, index+1).value = value
        row_index += 1

    # save the workbook
    wb.save(file_name)

if __name__ == "__main__":
    from load_non_standardized_settlements import load_non_standardized_settlements

    settlements_list = load_non_standardized_settlements(
        './datos/settlements/cuadro_to_use.xlsx'
    )

    save_standardized_settlements(settlements_list, './datos/exports/liquidaciones_estandarizadas.xlsx')