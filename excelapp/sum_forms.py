from django import forms
from django.forms.formsets import formset_factory

class ExcelForm(forms.Form):
    question = forms.CharField(max_length=256)
    answer = forms.CharField(max_length=100)

ExcelSumFormSet = formset_factory(ExcelForm, extra=5)