import pandas as pd

df = pd.read_excel('BD_roturas_final.xlsx')

def realizar_consulta(descfamilia, fecha_desde, fecha_hasta):
    if descfamilia and fecha_desde and fecha_hasta:
        fecha_desde = pd.Timestamp(fecha_desde)
        fecha_hasta = pd.Timestamp(fecha_hasta)

        resultado = df.loc[(df['descfamilia'] == descfamilia) & (df['FECHA'] >= fecha_desde) & (df['FECHA'] <= fecha_hasta)]
        
        total_bultos = resultado['bultos'].sum()
        total_costo = resultado['COSTO'].sum()
        unique_codart = len(resultado['codart'].unique().tolist())
        
        return total_bultos, total_costo, unique_codart

    return None, None, None