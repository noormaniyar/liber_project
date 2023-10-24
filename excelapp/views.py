# excelapp/views.py
from django.shortcuts import render
from django.http import HttpResponse
from openpyxl import Workbook
from .forms import ExcelFormSet
from openpyxl.utils import get_column_letter


def create_excel(request):
    if request.method == 'POST':
        formset = ExcelFormSet(request.POST)

        if formset.is_valid():
            # Create an Excel workbook
            workbook = Workbook()
            sheet = workbook.active

            # Write headers
            headers = ['Name', 'Age', 'City']
            for col_num, header in enumerate(headers):
                sheet.cell(row=1, column=col_num + 1, value=header)

            # Write data to the sheet and adjust column widths
            for row_num, form in enumerate(formset, start=2):
                data = [form.cleaned_data['name'], form.cleaned_data['age'], form.cleaned_data['city']]
                for col_num, value in enumerate(data):
                    cell = sheet.cell(row=row_num, column=col_num + 1, value=value)
                    column_letter = get_column_letter(col_num + 1)
                    column_width = len(str(value)) + 2
                    if sheet.column_dimensions[column_letter].width is None or sheet.column_dimensions[column_letter].width < column_width:
                        sheet.column_dimensions[column_letter].width = column_width  # Adjust column width

            # Set row heights based on content
            for row in sheet.iter_rows(min_row=2, max_row=len(formset) + 1):
                max_height = max(
                    cell.alignment.vertical if cell.alignment.vertical is not None else 0
                    for cell in row
                )
                sheet.row_dimensions[row[0].row].height = max_height

            # Create an HTTP response for the Excel file
            response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = 'attachment; filename="example.xlsx"'
            workbook.save(response)

            return response
    else:
        formset = ExcelFormSet()

    return render(request, 'excelapp/create_excel.html', {'formset': formset})
