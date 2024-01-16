from django.contrib import admin
from django.urls import path, include

from excelapp.views import generar_excel,obtener_clientes,obtener_productos, hello

urlpatterns = [
   path('', hello, name='generar_excel'),
   path('generar_excel', generar_excel, name='generar_excel'),
   path('obtener_clientes', obtener_clientes, name='obtener_clientes'),
   path('obtener_productos', obtener_productos, name='obtener_productos'),
]
