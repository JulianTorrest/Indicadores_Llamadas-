import pandas as pd
import json
import os

def obtener_diccionario_campos(ruta_archivo):
    """
    Lee un archivo Excel y devuelve un diccionario donde las llaves son 
    los nombres de las hojas y los valores son las listas de columnas.
    """
    try:
        # Cargamos el archivo Excel
        # Usamos pd.ExcelFile para evitar leer todos los datos de golpe y ahorrar memoria
        xls = pd.ExcelFile(ruta_archivo)
        
        diccionario_campos = {}
        
        for nombre_hoja in xls.sheet_names:
            # Leemos solo la primera fila (encabezados) para mayor eficiencia
            df = pd.read_excel(xls, sheet_name=nombre_hoja, nrows=0)
            diccionario_campos[nombre_hoja] = df.columns.tolist()
            
        return diccionario_campos

    except Exception as e:
        return f"Error al procesar el archivo: {e}"

# Configuración de la ruta
# Para Streamlit Cloud, asumiendo que el archivo está en la raíz del repo:
nombre_archivo = "Seguimiento encuestas consolidado.xlsx"


if os.path.exists(nombre_archivo):
    campos_hojas = obtener_diccionario_campos(nombre_archivo)
    
    # Guardar el diccionario en un archivo JSON para persistencia
    with open('estructura_campos.json', 'w', encoding='utf-8') as f:
        json.dump(campos_hojas, f, ensure_ascii=False, indent=4)
    
    print("Diccionario de campos generado y guardado en 'estructura_campos.json'")
    
    # Ejemplo de cómo se ve el diccionario
    for hoja, columnas in campos_hojas.items():
        print(f"Hoja: {hoja} | Cantidad de columnas: {len(columnas)}")
else:
    print(f"El archivo {nombre_archivo} no fue encontrado.")
