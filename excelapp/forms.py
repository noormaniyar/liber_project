from django import forms
from django.forms.formsets import formset_factory

class ExcelForm(forms.Form):
    name = forms.CharField(max_length=100)
    age = forms.IntegerField()
    city = forms.CharField(max_length=100)

ExcelFormSet = formset_factory(ExcelForm, extra=5)