import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

def convertir_columnas(df):
    df['Importe'] = df['Importe'].replace({r'[^\d.-]': ''}, regex=True).astype(float)
    df['Saldo'] = df['Saldo'].replace({r'[^\d.-]': ''}, regex=True).astype(float)
    return df

def extraer_info_adicional(df):
    df['CUIT'] = df['Info Adicional'].str.extract(r'CUIT:\s*(\d+)')
    df['Descripción'] = df['Info Adicional'].str.extract(r'Denominación:\s*(.*)')
    df.drop(columns=['Info Adicional'], inplace=True)
    return df

def modificar_excel(file, file_type):
    if file_type == 'xlsx':
        df = pd.read_excel(file, engine='openpyxl')
    else:
        df = pd.read_excel(file, engine='xlrd')
    
    df = convertir_columnas(df)
    df = extraer_info_adicional(df)
    return df

def procesar_pagos(file):
    df = pd.read_excel(file)

    # Convertir las columnas de fechas a datetime
    df['Fecha de vencimiento'] = pd.to_datetime(df['Fecha de vencimiento'])
    df['Fecha de Factura/Recibo'] = pd.to_datetime(df['Fecha de Factura/Recibo'])

    # Obtener la fecha de hoy
    today = datetime.today()

    # Crear columnas adicionales
    df['Semana'] = df['Fecha de vencimiento'].dt.to_period('W').apply(lambda r: r.start_time)
    df['Semana del Año'] = df['Fecha de vencimiento'].dt.isocalendar().week
    df['Día de la Semana'] = df['Fecha de vencimiento'].dt.day_name()

    # Crear la columna 'Estado de Deuda'
    df['Estado de Deuda'] = df.apply(lambda row: 'Deuda vencida' if row['Fecha de vencimiento'] < today and row['Estado de pago'] == 'No pagadas' else ('Deuda por pagar' if row['Estado de pago'] == 'No pagadas' else row['Estado de pago']), axis=1)

    # Dividir en pagados y otros
    df_pagados = df[df['Estado de pago'] == 'Pagado']
    df_otros = df[df['Estado de pago'] != 'Pagado']

    return df_pagados, df_otros

def procesar_clientes(file):
    df = pd.read_excel(file)

    # Convertir las columnas de fechas a datetime
    df['Fecha de vencimiento'] = pd.to_datetime(df['Fecha de vencimiento'])
    df['Fecha de Factura/Recibo'] = pd.to_datetime(df['Fecha de Factura/Recibo'])

    # Obtener la fecha de hoy
    today = datetime.today()

    # Crear columnas adicionales
    df['Semana'] = df['Fecha de vencimiento'].dt.to_period('W').apply(lambda r: r.start_time)
    df['Semana del Año'] = df['Fecha de vencimiento'].dt.isocalendar().week
    df['Día de la Semana'] = df['Fecha de vencimiento'].dt.day_name()

    # Crear la columna 'Estado de Cobranza'
    df['Estado de Cobranza'] = df.apply(lambda row: 'Cobranza vencida' if row['Fecha de vencimiento'] < today and row['Estado de pago'] == 'No pagadas' else ('Cobranza por cobrar' if row['Estado de pago'] == 'No pagadas' else row['Estado de pago']), axis=1)

    # Dividir en pagados y otros
    df_pagados = df[df['Estado de pago'] == 'Pagado']
    df_otros = df[df['Estado de pago'] != 'Pagado']

    return df_pagados, df_otros

st.title("Modificador de Archivos Excel")

uploaded_file_extracto = st.file_uploader("Subí el último extracto bancario", type=["xls", "xlsx"])
uploaded_file_pagos = st.file_uploader("Subí facturas de Proveedores", type=["xls", "xlsx"])
uploaded_file_clientes = st.file_uploader("Subí factura de Clientes", type=["xls", "xlsx"])

if uploaded_file_extracto is not None:
    file_type_extracto = uploaded_file_extracto.name.split('.')[-1]
    df_modificado = modificar_excel(uploaded_file_extracto, file_type_extracto)
    
    st.write("Vista previa del archivo de extracto bancario modificado:")
    st.dataframe(df_modificado.head(4))  # Mostrar solo las primeras 4 filas

    output_extracto = BytesIO()
    with pd.ExcelWriter(output_extracto, engine='xlsxwriter') as writer:
        df_modificado.to_excel(writer, index=False)
    output_extracto.seek(0)
    
    st.download_button(
        label="Descargar archivo de extracto bancario modificado",
        data=output_extracto,
        file_name="extracto_mod.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if uploaded_file_pagos is not None:
    df_pagados, df_otros = procesar_pagos(uploaded_file_pagos)
    
    st.write("Vista previa del archivo proveedores pagados:")
    st.dataframe(df_pagados.head(4))  # Mostrar solo las primeras 4 filas
    
    st.write("Vista previa de proveedores no pagados:")
    st.dataframe(df_otros.head(4))  # Mostrar solo las primeras 4 filas
    
    output_pagos = BytesIO()
    with pd.ExcelWriter(output_pagos, engine='xlsxwriter') as writer:
        df_pagados.to_excel(writer, sheet_name='deudas_pagadas', index=False)
        df_otros.to_excel(writer, sheet_name='deudas_no_pagadas', index=False)
    output_pagos.seek(0)
    
    st.download_button(
        label="Descargar archivo de deuda proveedores",
        data=output_pagos,
        file_name="proveedores_mod.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if uploaded_file_clientes is not None:
    df_pagados, df_otros = procesar_clientes(uploaded_file_clientes)
    
    st.write("Vista previa del archivo de clientes (Cobrados):")
    st.dataframe(df_pagados.head(4))  # Mostrar solo las primeras 4 filas
    
    st.write("Vista previa del archivo de clientes (No cobrados):")
    st.dataframe(df_otros.head(4))  # Mostrar solo las primeras 4 filas
    
    output_clientes = BytesIO()
    with pd.ExcelWriter(output_clientes, engine='xlsxwriter') as writer:
        df_pagados.to_excel(writer, sheet_name='cobranzas_cobradas', index=False)
        df_otros.to_excel(writer, sheet_name='cobranzas_no_cobradas', index=False)
    output_clientes.seek(0)
    
    st.download_button(
        label="Descargar archivo de clientes",
        data=output_clientes,
        file_name="clientes_mod.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
