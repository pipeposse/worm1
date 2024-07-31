import pandas as pd

# Crear un rango de fechas desde 2023 hasta 2030
start_date = '2023-01-01'
end_date = '2030-12-31'
date_range = pd.date_range(start=start_date, end=end_date)

# Crear un DataFrame con las fechas
df = pd.DataFrame(date_range, columns=['Fecha'])

# Agregar columnas de semana del año, día de la semana y trimestre
df['Semana del Año'] = df['Fecha'].dt.isocalendar().week
df['Día de la Semana'] = df['Fecha'].dt.day_name()
df['Trimestre'] = df['Fecha'].dt.to_period('Q').apply(lambda r: r.strftime('Q%q'))

# Guardar el DataFrame en un archivo Excel
file_path = 'calendario_2023_2030.xlsx'
df.to_excel(file_path, index=False)

print(f"Archivo guardado en: {file_path}")
