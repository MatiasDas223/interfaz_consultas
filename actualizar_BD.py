import os
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

def actualizar_BD():
    ruta_principal = r"\\layla\\Documentos\\STOCK"
    ruta_db = "BD_roturas.xlsx"

    # Verificar si la base de datos existe y obtener la última fecha registrada
    if os.path.exists(ruta_db):
        df_db = pd.read_excel(ruta_db)
        fecha_max = pd.to_datetime(df_db['FECHA'], dayfirst=True).max()
    else:
        fecha_max = None

    carpetas_anio = [nombre for nombre in os.listdir(ruta_principal) if os.path.isdir(os.path.join(ruta_principal, nombre)) and nombre.startswith("Stock")]

    # Si existe una fecha máxima, filtrar los años
    if fecha_max:
        year_max = fecha_max.year
        carpetas_anio = [carpeta for carpeta in carpetas_anio if int(carpeta[-4:]) >= year_max]

    datos_archivos = []

    for carpeta_anio in carpetas_anio:
        ruta_anio = os.path.join(ruta_principal, carpeta_anio, '10- ROTURAS', 'INFORME')
        carpetas_mes = [nombre for nombre in os.listdir(ruta_anio) if os.path.isdir(os.path.join(ruta_anio, nombre))]
        
        if fecha_max:
            month_max = fecha_max.month
            carpetas_mes = [carpeta for carpeta in carpetas_mes if int(carpeta.split('-')[0]) >= month_max]

        for carpeta_mes in carpetas_mes:
            ruta_mes = os.path.join(ruta_anio, carpeta_mes)
            archivos = [nombre for nombre in os.listdir(ruta_mes) if os.path.isfile(os.path.join(ruta_mes, nombre))]

            if archivos:
                archivo_reciente = max(archivos, key=lambda archivo: os.path.getmtime(os.path.join(ruta_mes, archivo)))
                ruta_archivo = os.path.join(ruta_mes, archivo_reciente)
                datos = pd.read_excel(ruta_archivo, sheet_name="Datos")
                
                # Si ya existen datos en la DB, filtrar solo los nuevos registros
                if fecha_max:
                    datos = datos[pd.to_datetime(datos['FECHA'], dayfirst=True) > fecha_max]

                datos_archivos.append(datos)
                print(f"Último archivo modificado en {carpeta_anio}/{carpeta_mes}: {archivo_reciente}")

    # Concatenar datos nuevos y guardar en la base de datos
    if datos_archivos:
        df_new_data = pd.concat(datos_archivos, ignore_index=True)
        if os.path.exists(ruta_db):
            df_db = pd.concat([df_db, df_new_data], ignore_index=True)
        else:
            df_db = df_new_data
        df_db.to_excel(ruta_db)

    # Nos quedamos con las columnas de interes

    df_db = df_db[['FECHA', 'COSTO' ,'CÓDIGO', 'DESCRIPCIÓN', 'Cantidad [Uni]', 'TIPO']]

    # Cambiar tipo de dato de la columna 'FECHA' a datetime con formato "%d/%m/%Y"
    df_db['FECHA'] = pd.to_datetime(df_db['FECHA'], format="%d/%m/%Y")

    # Cambiar tipo de dato de la columna 'DESCRIPCIÓN' a string
    df_db['DESCRIPCIÓN'] = df_db['DESCRIPCIÓN'].astype(str)

    maestro = pd.read_excel(r'H:\\STOCK\\MAESTRO ARTICULOS\\Maestro UdxBultoprov2023.xlsm')

    df_db.rename(columns= {'CÓDIGO' : 'codart'}, inplace=True)

    columnas_deseadas = ['codart', 'unidxbult', 'codfamilia', 'descfamilia', 'proveedor']
    maestro = maestro.loc[:, columnas_deseadas]
    maestro.dropna(inplace= True)

    maestro['codart']=maestro['codart'].astype(object)

    df = df_db.merge(maestro, on= 'codart')

    df['bultos'] = df['Cantidad [Uni]']/df['unidxbult']

    df.to_excel('BD_roturas_final.xlsx')

