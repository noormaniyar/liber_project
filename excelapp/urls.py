# excel_project/urls.py
from django.contrib import admin
from . import views
from django.urls import path


urlpatterns = [
    path('admin/', admin.site.urls),
    path('create-excel/', views.create_excel, name='create_excel'),
    path('create-pdf/', views.create_pdf, name='create_pdf'),
    path('create-sum-pdf/', views.create_sum_pdf, name='create_sum_pdf'),
    path('generate-pdf/', views.generate_pdf, name='generate_pdf'),
    
]
