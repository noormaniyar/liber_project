from django import forms
from django.forms.formsets import formset_factory
from .models import TableConfig




class TableConfigForm(forms.ModelForm):
    class Meta:
        model = TableConfig
        fields = ['num_rows', 'num_cols', 'cell_width', 'cell_height']
        # fields = ['num_rows', 'num_cols', 'cell_width', 'cell_height', 'merge_cells']
        



class ExcelForm(forms.Form):
    name = forms.CharField(max_length=100)
    age = forms.IntegerField()
    city = forms.CharField(max_length=100)

ExcelFormSet = formset_factory(ExcelForm, extra=2)