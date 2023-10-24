from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from io import BytesIO




# User input for number of rows and columns
num_rows = 5  # Replace with user input
num_cols = 4  # Replace with user input

# User input for cell width and height in inches
cell_width = 1.0  # Replace with user input
cell_height = 0.5  # Replace with user input

# Create a PDF document
pdf_buffer = BytesIO()
doc = SimpleDocTemplate(pdf_buffer, pagesize=(landscape(letter)))

# Create an empty table with the specified number of rows and columns
data = [['' for _ in range(num_cols)] for _ in range(num_rows)]
table = Table(data, colWidths=[cell_width * inch] * num_cols, rowHeights=[cell_height * inch] * num_rows)

# Set custom style for the table
table.setStyle(TableStyle([
    ('INNERGRID', (0, 0), (-1, -1), 0.5, (0, 0, 0)),  # Add inner grid lines
    ('BOX', (0, 0), (-1, -1), 0.5, (0, 0, 0))  # Add cell borders
]))

# Build the PDF
elements = [table]
doc.build(elements)

# Save or serve the PDF as needed
pdf_buffer.seek(0)
with open('blank_table.pdf', 'wb') as f:
    f.write(pdf_buffer.read())
