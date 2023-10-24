# excelapp/views.py
from django.shortcuts import render
from django.http import HttpResponse
from openpyxl import Workbook
from .forms import ExcelForm


def create_excel(request):
    if request.method == 'POST':
        form = ExcelForm(request.POST)

        if form.is_valid():
            name = form.cleaned_data['name']
            age = form.cleaned_data['age']

            # Create an Excel workbook
            workbook = Workbook()
            sheet = workbook.active

            # Write data to the sheet
            sheet['A1'] = 'Name'
            sheet['B1'] = 'Age'
            sheet['A2'] = name
            sheet['B2'] = age

            # Create an HTTP response for the Excel file
            response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = 'attachment; filename="example.xlsx"'
            workbook.save(response)

            return response
    else:
        form = ExcelForm()

    return render(request, 'excelapp/create_excel.html', {'form': form})
