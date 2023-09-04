def generar_informe_mensual(df):
    import os
    import pandas as pd
    from datetime import datetime
    from openpyxl import load_workbook
    
    # Obtener el mes y año actual
    current_month = datetime.now().month
    current_year = datetime.now().year

    # Filtrar para excluir registros del mes actual
    df = df[~((df['FECHA'].dt.month == current_month) & 
                    (df['FECHA'].dt.year == current_year))]
    
    # Ahora df_roturas tiene todos los registros excepto aquellos del mes actual

    # Convertimos la columna 'FECHA' a formato de fecha
    df_pivot= df.copy() # Le asignamos el .copy para que las modificaciones a df_pivot no afecten a df, sino quedan vinculados
    df_pivot.dropna(inplace= True)
    df_pivot['FECHA'] = pd.to_datetime(df_pivot['FECHA'])

    # Creamos una nueva columna 'AÑO' extrayendo el año de la fecha
    df_pivot['AÑO'] = df_pivot['FECHA'].dt.year

    # Creamos una nueva columna 'MES' extrayendo el mes de la fecha
    df_pivot['MES'] = df_pivot['FECHA'].dt.month

    # Re-mapeamos los valores en la columna 'TIPO'
    mapeo_tipos = {
        'ROTURA': 'ROTURA',
        'ROTURA DEPOSITO': 'ROTURA',
        'VENCIMIENTO': 'VENCIMIENTO',
        'VENCIMIENTO DEPOSITO': 'VENCIMIENTO'
    }
    df_pivot['TIPO'] = df_pivot['TIPO'].replace(mapeo_tipos)

    #df para la pivot
    tipos = ['ROTURA', 'VENCIMIENTO']
    df_pivot = df_pivot[df_pivot['TIPO'].isin(tipos)]

    # Creamos la tabla pivote
    pivot = pd.pivot_table(df_pivot, values='bultos', index=['MES', 'TIPO'], columns='AÑO', aggfunc='sum').reset_index()

    mes = (datetime.now().month) # Que nos muestre solo a partir del mes actual
    pivot.fillna(0, inplace= True) # Los nulos los cambiamos por 0
    pivot = pivot[pivot['MES'] < mes] # que nos muestre solo a partir del mes actual
    columnas = pivot.columns #Conseguimos los nombres de las columnas
    pivot['RELACION %']=((pivot[columnas[len(columnas)-1]]-pivot[columnas[len(columnas)-2]])/pivot[columnas[len(columnas)-1]])*100
    
    pivot.iloc[:, -1] = pivot.iloc[:, -1].apply(lambda x: "{:,.2f}%".format(x))
    pivot[columnas[len(columnas)-1]] = pivot[columnas[len(columnas)-1]].apply(lambda x: "{:,.2f} Bul.".format(x))
    pivot[columnas[len(columnas)-2]] = pivot[columnas[len(columnas)-2]].apply(lambda x: "{:,.2f} Bul.".format(x))
    pivot[columnas[len(columnas)-3]] = pivot[columnas[len(columnas)-3]].apply(lambda x: "{:,.2f} Bul.".format(x))
    pivot[columnas[len(columnas)-4]] = pivot[columnas[len(columnas)-4]].apply(lambda x: "{:,.2f} Bul.".format(x))
    
    # obteniendo el año y mes actual
    current_year = datetime.now().year
    current_month = datetime.now().month

    # obteniendo el mes y año anterior
    if current_month == 1:
        last_month = 12
        last_year = current_year - 1
    else:
        last_month = current_month - 1
        last_year = current_year

    # Filtrando el dataframe para incluir solo las filas del mes anterior
    df_last_month = df[(df['FECHA'].dt.year == last_year) & (df['FECHA'].dt.month == last_month)].copy()
    df_last_month.dropna(inplace= True)
    
    # Convertir los valores en la columna 'COSTO' a números
    #df['COSTO'] = pd.to_numeric(df['COSTO'], errors='coerce')

    agregacion = df_last_month.groupby('TIPO')['COSTO'].sum().astype(int).reset_index().sort_values(by= 'COSTO', ascending= False) # Agrupamos por tipo sumando el costo

    total_costo = agregacion['COSTO'].sum() # Creamos la variable total que es la suma de todos los costos

    total_row = pd.DataFrame({'TIPO': ['TOTAL'], 'COSTO': [total_costo]})# Creamos la fila total

    agregacion = pd.concat([agregacion, total_row], ignore_index=True)# Concatenamos la fila TOTAL

    agregacion['Proporcion'] = (agregacion['COSTO']/total_costo)*100

    agregacion['COSTO'] = agregacion['COSTO'].apply(lambda x: "${:,.0f}".format(x)) # Agregamos formato de pesos

    agregacion['Proporcion'] = agregacion['Proporcion'].apply(lambda x: "{:,.2f}%".format(x)) # Agregamos formato de porcentaje
        
    agregacion_2 = df_last_month.groupby(['descfamilia', 'TIPO'])['COSTO'].sum().astype(int).reset_index() # Agrupamos por tipo sumando el costo

    agregacion_2.sort_values(by= 'COSTO', ascending= False, inplace= True) #Ordenamos por costo

    agregacion_2.reset_index(drop= True,inplace= True) #Reiniciamos los indices

    agregacion_2 = agregacion_2.loc[agregacion_2['COSTO'] != 0]

    tablas = agregacion_2['TIPO'].unique()
    
    agregacion_2.rename(columns= {'descfamilia' : 'FAMILIA'}, inplace= True)
    
    lista_tablas = []

    for x in tablas:
        lista = agregacion_2[agregacion_2['TIPO']== x].loc[:, ('FAMILIA', 'COSTO')]
        lista['Proporcion del Area'] = (100*(lista['COSTO']/lista['COSTO'].sum())).apply(lambda x: "{:,.2f}%".format(x))
        lista['Proporcion del Total'] = (100*(lista['COSTO']/total_costo)).apply(lambda x: "{:,.2f}%".format(x))
        lista['COSTO'] = lista['COSTO'].apply(lambda x: "${:,.2f}".format(x)) # Agregamos formato de pesos
        lista_tablas.append(lista)

    # Definir los tipos de roturas y vencimientos
    roturas_tipo = ('ROTURA', 'ROTURA DEPOSITO')
    vencimientos_tipo = ('VENCIMIENTO', 'VENCIMIENTO SALON')

    # Filtrar y realizar la agregación por roturas
    agregacion_roturas = df_last_month[df_last_month['TIPO'].isin(roturas_tipo)]
    agregacion_roturas = agregacion_roturas.groupby('proveedor')['COSTO'].sum().reset_index()

    # Filtrar y realizar la agregación por vencimientos
    agregacion_vencimientos = df_last_month[df_last_month['TIPO'].isin(vencimientos_tipo)]
    agregacion_vencimientos = agregacion_vencimientos.groupby('proveedor')['COSTO'].sum().reset_index()

    # Combinar los resultados de roturas y vencimientos
    resultado_final = pd.merge(agregacion_roturas, agregacion_vencimientos, on='proveedor', how='outer')
    resultado_final.columns = ['Proveedor', 'Rotura', 'Vencimiento']
    resultado_final.fillna(0,inplace= True)
    resultado_final['Total'] = resultado_final['Rotura']+resultado_final['Vencimiento']
    resultado_final.sort_values(by = 'Total', inplace=True, ascending= False)
    resultado_final['Proporcion_Total'] = (resultado_final['Total'] / total_costo)*100

    #Formateamos los valores

    resultado_final['Rotura'] = resultado_final['Rotura'].apply(lambda x: "${:,.2f}".format(x))
    resultado_final['Vencimiento'] = resultado_final['Vencimiento'].apply(lambda x: "${:,.2f}".format(x))
    resultado_final['Total'] = resultado_final['Total'].apply(lambda x: "${:,.2f}".format(x))
    resultado_final['Proporcion_Total'] = resultado_final['Proporcion_Total'].apply(lambda x: "{:,.2f}%".format(x))

    from reportlab.lib.styles import getSampleStyleSheet

    from reportlab.lib.styles import ParagraphStyle

    # Crear un nuevo estilo para el subtítulo a partir del estilo 'Title'
    estilos = getSampleStyleSheet()
    estilo_subtitulo = estilos['Title'].clone('subtitulo')  # Crea una copia del estilo 'Title'
    estilo_subtitulo.fontSize = 14  # Cambiar el tamaño de la fuente
    estilo_subtitulo.leading = 16   # Cambiar el espaciado entre líneas

    # Definir un nuevo estilo para los títulos de las tablas
    estilo_titulo_tabla = ParagraphStyle(
        'TituloTabla',  # Nombre del estilo
        parent=estilos['BodyText'],  # Estilo base
        alignment=1,  # 1 = centrado, 0 = izquierda, 2 = derecha
        fontSize=14,  # Tamaño de la fuente
        spaceAfter=12,  # Espacio después del título
)

    # Diccionario para traducir los nombres de los meses al español
    meses_en_espanol = [
        'Enero',
        'Febrero',
        'Marzo',
        'Abril',
        'Mayo',
        'Junio',
        'Julio',
        'Agosto',
        'Septiembre',
        'Octubre',
        'Noviembre',
        'Diciembre',
    ]

    from dateutil.relativedelta import relativedelta

    # Obtener la fecha actual y restarle un mes
    fecha_anterior = datetime.now() - relativedelta(months=1)

    # Obtener el índice del mes anterior (donde enero es 0)
    indice_mes_anterior = fecha_anterior.month - 1

    # Obtener el año del mes anterior
    ano_anterior = fecha_anterior.year

    # Traducir el mes al español
    mes_anterior = meses_en_espanol[indice_mes_anterior]

    ruta_principal = r"\\layla\\Documentos\\STOCK"

    carpetas_anio = [nombre for nombre in os.listdir(ruta_principal) if os.path.isdir(os.path.join(ruta_principal, nombre)) and nombre.startswith("Stock")]

    # Ordenar las carpetas basado en el número al final del nombre
    carpetas_anio_sorted = sorted(carpetas_anio, key=lambda x: int(x[len("Stock"):]))

    # Tomar la última carpeta
    carpeta_anio = carpetas_anio_sorted[-1] if carpetas_anio_sorted else None

    # Ruta de la carpeta que deseas crear
    ruta_carpeta = r"H:\\STOCK\\"

    ruta_carpeta = os.path.join(ruta_carpeta, carpeta_anio)

    ruta_carpeta = os.path.join(ruta_carpeta, f'10- ROTURAS\\INFORME MENSUAL\\{mes_anterior}')

    # Verificar si la carpeta no existe y crearla si es necesario
    if not os.path.exists(ruta_carpeta):
        os.makedirs(ruta_carpeta)

    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Frame
    from reportlab.lib import colors
    from datetime import date
    from reportlab.platypus import PageBreak
    from reportlab.lib.units import inch, mm
    from reportlab.platypus import Image
    from reportlab.pdfgen import canvas

    # Obtener la fecha actual
    fecha_elaboracion = date.today().strftime("%d/%m/%Y")

    # Usar el mes anterior en el subtítulo
    subtitulo = f"Análisis del mes de {mes_anterior} del {ano_anterior}"

    # Convertir el DataFrame en una lista de listas
    data_tabla1 = [agregacion.columns.tolist()] + agregacion.values.tolist()

    # Crear el archivo PDF dentro de la carpeta
    nombre_archivo = f"Reporte Mensual Roturas {mes_anterior} del {ano_anterior}.pdf"
    ruta_archivo = os.path.join(ruta_carpeta, nombre_archivo)
    doc = SimpleDocTemplate(ruta_archivo, pagesize=letter)
    elementos = []

    # Título y espaciado
    titulo = "Reporte de bajas por rotura mensual"
    estilos = getSampleStyleSheet()
    elementos.append(Paragraph(titulo, estilos['Title']))

    #### TABLA COMPARATIVA ENTRE MESES DE LAS BAJAS

    # Agregar el subtítulo a los elementos
    elementos.append(Paragraph(subtitulo, estilo_subtitulo))
    elementos.append(Spacer(1, 12))  # Espacio entre el subtítulo y el texto adicional

    # Espaciado después del título
    elementos.append(Spacer(1, 12))

    # Texto adicional
    texto_adicional = "En la siguiente tabla se presenta una comparacion en bultos de la cantiad de bajas vs el mismo mes del año anterior."
    elementos.append(Paragraph(texto_adicional, estilo_titulo_tabla))  # Agregando el texto adicional

    # Espaciado después del texto adicional
    elementos.append(Spacer(1, 24))

    # Crear la Tabla 1 y agregar a los elementos
    tabla1 = Table([pivot.columns.tolist()] + pivot.values.tolist())
    tabla1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elementos.append(tabla1)  # Agregar la tabla a los elementos

    elementos.append(Spacer(1, 24))
    elementos.append(PageBreak())

    ##### TABLA RESUMEN

    # Texto adicional
    texto_adicional = "En la siguiente tabla se presenta las distintas causas de las bajas con su monto relacionado."
    elementos.append(Paragraph(texto_adicional, estilo_titulo_tabla))  # Agregando el texto adicional

    # Espaciado después del texto adicional
    elementos.append(Spacer(1, 24))

    # Crear la Tabla 1 y agregar a los elementos
    tabla2 = Table(data_tabla1)
    tabla2.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elementos.append(tabla2)  # Agregar la tabla a los elementos

    elementos.append(Spacer(1, 24))
    elementos.append(PageBreak())


    #### TABLAS POR FAMILIA Y COSTO DE LAS BAJAS


    # Generar las tablas y agregarlas al PDF
    for i, tabla in enumerate(lista_tablas):
        # Crear el título de la tabla
        titulo_tabla = f"Tabla de costos de bajas por {tablas[i]} desagregado por familia de productos"
        elementos.append(Paragraph(titulo_tabla, estilo_titulo_tabla))

        # Obtener los encabezados y los datos de la tabla
        encabezados = tabla.columns.tolist()
        datos = tabla.values.tolist()

        # Crear la tabla
        tabla_pdf = Table([encabezados] + datos)
        tabla_pdf.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        # Agregar la tabla al PDF
        elementos.append(tabla_pdf)
        elementos.append(PageBreak())  # Agregar un salto de página después de cada tabla


    #### TABLA DE PROVEEDORES

    texto_adicional = "En la siguiente tabla los costos asociados a las bajas por proveedor."
    elementos.append(Paragraph(texto_adicional, estilo_titulo_tabla))  # Agregando el texto adicional


    # Espaciado después del texto adicional
    elementos.append(Spacer(1, 24))

    # Crear la Tabla 1 y agregar a los elementos
    tabla2 = Table([resultado_final.columns.tolist()] + resultado_final.values.tolist())
    tabla2.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elementos.append(tabla2)  # Agregar la tabla a los elementos

    elementos.append(Spacer(1, 24))
    elementos.append(PageBreak())


    #### DARLE FORMATO AL PDF

    class NumberedCanvas(canvas.Canvas):
        def __init__(self, *args, **kwargs):
            canvas.Canvas.__init__(self, *args, **kwargs)
            self._saved_page_states = []

        def showPage(self):
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            """add page info to each page (page x of y)"""
            num_pages = len(self._saved_page_states)
            for state in self._saved_page_states:
                self.__dict__.update(state)
                self.setFont("Helvetica", 10)
                self.drawRightString(600, 30, f"Página {self._pageNumber} of {num_pages}")
                canvas.Canvas.showPage(self)
            canvas.Canvas.save(self)
        
    def agregar_encabezado(canvas, doc):
        """
        Agrega la fecha de elaboración y la leyenda en la parte superior de cada página,
        y una imagen en la primera página.
        """
        canvas.saveState()
        canvas.setFont("Helvetica", 10)
        canvas.drawString(450, doc.height + doc.topMargin +10 , f"Fecha de Elaboración: {fecha_elaboracion}")
        canvas.drawString(450, doc.height + doc.topMargin +20, "Elaborado por Matias da Silva")  # Agregar la leyenda

        if doc.page == 1:
            img_path = '../Roturas/logo-17.png'  # Cambia esto a la ruta de tu imagen
            img = Image(img_path, width=2*inch, height=1*inch)  # Cambia el tamaño según lo que necesites
            img.drawOn(canvas, 30, doc.height + doc.topMargin - 10)

        canvas.restoreState()


    # Asociar las funciones al documento
    doc.build(elementos, 
            onFirstPage=agregar_encabezado, 
            onLaterPages=agregar_encabezado,
            canvasmaker=NumberedCanvas)

    # Guardar las tablas en formato Excel
    ruta_pivot = os.path.join(ruta_carpeta, "Comparacion Anual.xlsx")
    pivot.to_excel(ruta_pivot, index=False)

    ruta_agregacion_2 = os.path.join(ruta_carpeta, "Agregacion por familias.xlsx")
    agregacion_2.to_excel(ruta_agregacion_2, index=False)

    ruta_resultado_final = os.path.join(ruta_carpeta, "Agregacion por proveedores.xlsx")
    resultado_final.to_excel(ruta_resultado_final, index=False)

    pass
