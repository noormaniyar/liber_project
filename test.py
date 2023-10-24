from openpyxl import Workbook
from openpyxl.worksheet.dimensions import ColumnDimension, RowDimension

# Create a new workbook and select the active sheet
workbook = Workbook()
sheet = workbook.active

# Set the dimensions of columns A and B
column_A = sheet.column_dimensions['A']
column_A.width = 15  # Set the width of column A to 15

# Set the height of row 1
row_1 = sheet.row_dimensions[1]
row_1.height = 20  # Set the height of row 1 to 20

# Write data to cells
sheet['A1'] = 'Name'
sheet['B1'] = 'Age'
sheet['A2'] = 'Alice'
sheet['B2'] = 25

# Save the workbook
workbook.save('example.xlsx')
