# excel_project/urls.py
from django.contrib import admin
from . import views
from django.urls import path


urlpatterns = [
    path('admin/', admin.site.urls),
    path('create-excel/', views.create_excel, name='create_excel'),
    
]
