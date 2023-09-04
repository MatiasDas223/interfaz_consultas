import tkinter as tk
from tkinter import ttk
import pandas as pd
from actualizar_BD import actualizar_BD
from calculo_roturas import realizar_consulta
from informe import generar_informe_mensual

def actualizar_datos():
    global df
    df = pd.read_excel('BD_roturas_final.xlsx')

# Función para autocompletar guiones y convertir el formato de la fecha
def adaptar_fecha_formato(fecha):
    partes = fecha.split("-")
    if len(partes) == 3 and len(partes[0]) == 2 and len(partes[1]) == 2 and len(partes[2]) == 4:
        # Convertir de "DD-MM-AAAA" a "MM-DD-AAAA"
        return f"{partes[1]}-{partes[0]}-{partes[2]}"
    else:
        return None

def consulta_callback():
    descfamilia_input = descfamilia_var.get()
    fecha_desde_input = adaptar_fecha_formato(fecha_desde_var.get())
    fecha_hasta_input = adaptar_fecha_formato(fecha_hasta_var.get())
    
    if not fecha_desde_input or not fecha_hasta_input:
        # Manejo de errores en el formato de fecha ingresado
        total_bultos_var.set("Error en el formato de fecha")
        return

    total_bultos, total_costo, unique_codart = realizar_consulta(descfamilia_input, fecha_desde_input, fecha_hasta_input)
    
    total_bultos_var.set(f"Bultos totales: {round(total_bultos, 2) if total_bultos is not None else 'N/A'}")
    total_costo_var.set(f"Costo total: ${round(total_costo, 2) if total_costo is not None else 'N/A'}")
    unique_codart_var.set(f"Códigos únicos: {unique_codart if unique_codart is not None else 'N/A'}")

app = tk.Tk()
app.title("Consulta de Datos")

frame = ttk.Frame(app)
frame.grid(row=0, column=0, padx=10, pady=10)

descfamilia_label = ttk.Label(frame, text="Familia:")
descfamilia_label.grid(row=0, column=0, pady=5, sticky="w")
descfamilia_var = tk.StringVar()
descfamilia_entry = ttk.Entry(frame, textvariable=descfamilia_var)
descfamilia_entry.grid(row=0, column=1, pady=5, sticky="ew")

actualizar_datos_btn = ttk.Button(frame, text="Actualizar datos", command=actualizar_datos)
actualizar_datos_btn.grid(row=0, column=2, pady=5, padx=5)

fecha_desde_label = ttk.Label(frame, text="Fecha Desde (DD-MM-AAAA):")
fecha_desde_label.grid(row=1, column=0, pady=5, sticky="w")
fecha_desde_var = tk.StringVar()
fecha_desde_entry = ttk.Entry(frame, textvariable=fecha_desde_var)
fecha_desde_entry.grid(row=1, column=1, pady=5, sticky="ew")

fecha_hasta_label = ttk.Label(frame, text="Fecha Hasta (DD-MM-AAAA):")
fecha_hasta_label.grid(row=2, column=0, pady=5, sticky="w")
fecha_hasta_var = tk.StringVar()
fecha_hasta_entry = ttk.Entry(frame, textvariable=fecha_hasta_var)
fecha_hasta_entry.grid(row=2, column=1, pady=5, sticky="ew")

actualizar_btn = ttk.Button(frame, text="Actualizar BD", command=actualizar_BD)
actualizar_btn.grid(row=1, column=2, padx=5)

consulta_btn = ttk.Button(frame, text="Realizar Consulta", command=consulta_callback)
consulta_btn.grid(row=3, column=0, columnspan=2, pady=10)

informe_btn = ttk.Button(frame, text="Informe Mensual", command=lambda: generar_informe_mensual(df))
informe_btn.grid(row=3, column=2, columnspan=2, pady=10)

total_bultos_var = tk.StringVar()
total_bultos_label = ttk.Label(frame, textvariable=total_bultos_var)
total_bultos_label.grid(row=4, column=0, columnspan=2, pady=5)

total_costo_var = tk.StringVar()
total_costo_label = ttk.Label(frame, textvariable=total_costo_var)
total_costo_label.grid(row=5, column=0, columnspan=2, pady=5)

unique_codart_var = tk.StringVar()
unique_codart_label = ttk.Label(frame, textvariable=unique_codart_var)
unique_codart_label.grid(row=6, column=0, columnspan=2, pady=5)

app.mainloop()
