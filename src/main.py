import os
import pandas as pd

# Ruta del archivo Excel de entrada
input_filename = 'C:/Users/wppaez/Desktop/BASE DE DATOS/input/database.xlsx'

# Carga el archivo Excel en un DataFrame de pandas
try:
    df = pd.read_excel(input_filename)
except Exception as e:
    print(f'Error al cargar el archivo Excel: {str(e)}')
    exit(1)

# Nombre de la columna que deseas recorrer
columna = 'N° MESES'

# Recorre las filas y realiza la inserción
for index, row in df.iterrows():
    numero = row[columna]

    # Verifica si la celda tiene un valor numérico positivo mayor a 1
    if isinstance(numero, (int, float)) and numero > 1:
        # Calcula el número de filas a insertar como valor de la celda - 1
        num_filas_insertar = int(numero) - 1

        # Crea las filas a insertar con el mismo valor de la celda
        nuevas_filas = pd.DataFrame([row] * num_filas_insertar)

        # Encuentra la ubicación actual de la fila en el DataFrame
        idx = df.index.get_loc(index) 

        # Inserta las filas inmediatamente debajo
        df = pd.concat([df.iloc[:idx + 1], nuevas_filas, df.iloc[idx + 1:]], ignore_index=True)

        # Actualiza la columna "current month" en las filas recién insertadas
        for i in range(idx + 1, (idx + 1 + num_filas_insertar)):
            df.at[i, 'current month'] = i - (idx)

    # Actualiza la columna "current month" en la fila actual
    df.at[index, 'current month'] = 1

# Carpeta de salida
output_folder = 'output'
os.makedirs(output_folder, exist_ok=True)

# Ruta completa del archivo de salida
output_file_path = os.path.join(output_folder, 'archivo_modificado.xlsx')

# Guarda el DataFrame modificado en un nuevo archivo Excel
try:
    df.to_excel(output_file_path, index=False)
    print(f'Archivo modificado guardado en: {output_file_path}')
except Exception as e:
    print(f'Error al guardar el archivo modificado: {str(e)}')