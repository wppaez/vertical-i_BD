import os
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta

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
        for i in range(1, num_filas_insertar + 1):
            nueva_fila = dict(row)  # Copia los valores de la fila actual
            nueva_fila['current month'] = i  # Actualiza 'current month' en la nueva fila

            if i == num_filas_insertar:
                # La última fila insertada toma el valor original de la celda 'N° DIAS'
                nueva_fila['N° DIAS'] = row['N° DIAS']
            else:
                # Las filas anteriores toman el valor 0 en 'N° DIAS'
                nueva_fila['N° DIAS'] = 0

            df = pd.concat([df.iloc[:idx + 1], pd.DataFrame([nueva_fila]), df.iloc[idx + 1:]], ignore_index=True)

        # Actualiza la columna "current month" en las filas recién insertadas
        for i in range(idx + 1, (idx + 1 + num_filas_insertar)):
            df.at[i, 'current month'] = i - (idx)

    # Actualiza la columna "current month" en la fila actual
    df.at[index, 'current month'] = 1

# Reemplaza los valores NaN en 'current month' con el valor de la columna 'N° MESES'
df['current month'].fillna(df[columna], inplace=True)

# Reemplaza los valores NaN en 'N° DIAS' con 0
df['N° DIAS'].fillna(0, inplace=True)

# Establece la columna 'N° DIAS' en 0 para las filas insertadas
df.loc[df['current month'] > 1, 'N° DIAS'] = 0

# Convierte la columna 'current month' en enteros
df['current month'] = df['current month'].apply(lambda x: int(x) if not pd.isna(x) else x)

# Definir una función para aplicar la suma de meses y días a cada fila
def suma_meses_y_dias(row):
    fecha_inicial = row['FECHA DE INICIO']
    meses_a_sumar = row['current month']
    dias_a_sumar = row['N° DIAS']
    fecha_resultado = fecha_inicial + relativedelta(months=meses_a_sumar, days=dias_a_sumar)
    return fecha_resultado

# Aplicar la función a cada fila del DataFrame y crear una nueva columna 'current_date'
df['current_date'] = df.apply(suma_meses_y_dias, axis=1)

# Formatea las columnas de fecha
df['FECHA DE INICIO'] = df['FECHA DE INICIO'].dt.strftime('%d/%m/%Y')
df['FECHA DE FIN'] = df['FECHA DE FIN'].dt.strftime('%d/%m/%Y')
df['current_date'] = df['current_date'].dt.strftime('%d/%m/%Y')

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

