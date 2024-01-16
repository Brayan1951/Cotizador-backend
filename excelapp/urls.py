from django.contrib import admin
from django.urls import path, include

from excelapp.views import generar_excel, hello

urlpatterns = [
   path('', hello, name='generar_excel'),
   path('generar_excel', generar_excel, name='generar_excel'),
]
