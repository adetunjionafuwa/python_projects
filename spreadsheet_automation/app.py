import openpyxl as xl

#load workbook from directory
wb = xl.load_workbook('transactions.xlsx')

# Grab excel sheet to work with.
# In order to know the available sheet names,
# use the wb.sheetnames to view sheet arrays
sheet = wb['Sheet1']

# Access sheet cell
for row in range(2, sheet.max_row + 1):
    cell = sheet.cell(row, 3)
    corrected_price = cell.value *0.9
    corrected_price_cell = sheet.cell(row, 4)
    corrected_price_cell.value = corrected_price


wb.save('transactions.xlsx')