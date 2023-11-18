import openpyxl

def search_name_in_excel(file_path, target_name):
    # Load the Excel file
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Variables to store the last group and row index
    last_group = None
    last_row_index = None

    # Iterate through the rows in the sheet
    for idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        for cell_value in row:
            if "GROUP" in str(cell_value):
                last_group = cell_value
                last_row_index = idx

            if cell_value == target_name:
                return row, idx, last_group, last_row_index

    return None, None, None, None

# Specify the file path and the target name to search
file_path = '100LVL CSC.xlsx'
target_name = '23/12009'

# Call the function
found_row, row_index, last_group, last_row_index = search_name_in_excel(file_path, target_name)

if found_row is not None:
    print(f'The name {target_name} was found in row {row_index}.')
    for idx, cell_value in enumerate(found_row, start=1):
        print(f'Column {idx}: {cell_value}')
    if last_group is not None and last_row_index is not None:
        print(f'Group: {last_group}')
    else:
        print("No GROUP found before the target name.")
else:
    print(f'The name {target_name} was not found in the Excel file.')
