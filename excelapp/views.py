import json
from django.shortcuts import render
from django.http import HttpResponse
import pandas as pd

# Create your views here.
def hello(request):
    return HttpResponse("Holiwis world")

def generar_excel(request):
    # if request.method =='POST':
    #     print("estas en el post")
    #     codigo = request.POST.get('codigo')
    #     # codigo_sap = request.POST.get('codigo_sap')
    #     # descripcion = request.POST.get('descripcion')
    #     # return {codigo}
    #     return HttpResponse("hola")
        
    if request.method =='GET':
        print("pasaste el get")
        return HttpResponse("estas en el gett")
    else:
        data = json.loads(request.body.decode('utf-8'))
        print(data)
        df = pd.DataFrame(data)
        print("pasaste el post")
        archivo_temporal = "temp.xlsx"
        df.to_excel(archivo_temporal, index=False)

        response = HttpResponse(open(archivo_temporal, 'rb').read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename=temp.xlsx'
        return response
        
        
    