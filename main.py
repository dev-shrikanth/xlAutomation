from openpyxl import load_workbook
from utils.sheetfunctions import get_sheet_and_range, set_formulae
from utils.wbk import init_wb, finalize_wb
import shutil


# create a copy of 'output.xlsx' and name it 'output1.xlsx'
shutil.copy("output.xlsx", "output1.xlsx")

# Load and initialize the workbook
wb = load_workbook("output1.xlsx")
init_wb(wb)

# Populate Revenue sheet
backend, range_address = get_sheet_and_range(wb, "albums")
ws_backend = wb[backend]
ws_dest = wb["Revenue"]

for row in ws_backend[range_address]:
    # ignore the first title row
    if row[0].row != 1:
        ws_dest["B" + str(row[0].row)] = row[0].value
        ws_dest["C" + str(row[0].row)] = row[1].value
        ws_dest["E" + str(row[0].row)] = row[2].value
        # Break after 10 iterations
        if row[0].row == 10:
            break


# Set formulae for Revenue sheet
set_formulae(wb)

# Finalize the workbook
finalize_wb(wb)
wb.save("output1.xlsx")
