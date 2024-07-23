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

    # Filtrar las facturas que no están pagadas y las que están pagadas
    unpaid = df[df['Estado de pago'] != 'Pagado']
    paid = df[df['Estado de pago'] == 'Pagado']

    # Crear columnas adicionales para ambos dataframes
    for dataframe in [unpaid, paid]:
        dataframe['Semana'] = dataframe['Fecha de vencimiento'].dt.to_period('W').apply(lambda r: r.start_time)
        dataframe['Semana del Año'] = dataframe['Fecha de vencimiento'].dt.isocalendar().week
        dataframe['Día de la Semana'] = dataframe['Fecha de vencimiento'].dt.day_name()

    return unpaid, paid

def procesar_clientes(file):
    df = pd.read_excel(file)

    # Convertir las columnas de fechas a datetime
    df['Fecha de vencimiento'] = pd.to_datetime(df['Fecha de vencimiento'])
    df['Fecha de Factura/Recibo'] = pd.to_datetime(df['Fecha de Factura/Recibo'])

    # Obtener la fecha de hoy
    today = datetime.today()

    # Filtrar las facturas que no están pagadas y las que están pagadas
    unpaid = df[df['Estado de pago'] != 'Pagado']
    paid = df[df['Estado de pago'] == 'Pagado']

    # Crear columnas adicionales para ambos dataframes
    for dataframe in [unpaid, paid]:
        dataframe['Semana'] = dataframe['Fecha de vencimiento'].dt.to_period('W').apply(lambda r: r.start_time)
        dataframe['Semana del Año'] = dataframe['Fecha de vencimiento'].dt.isocalendar().week
        dataframe['Día de la Semana'] = dataframe['Fecha de vencimiento'].dt.day_name()

    return unpaid, paid

st.title("Modificador de Archivos Excel")

uploaded_file_extracto = st.file_uploader("Subí el último extracto bancario", type=["xls", "xlsx"])
uploaded_file_pagos = st.file_uploader("Subí el archivo de pagos (proveedores)", type=["xls", "xlsx"])
uploaded_file_clientes = st.file_uploader("Subí el archivo de clientes (cobros)", type=["xls", "xlsx"])

if uploaded_file_extracto is not None:
    file_type_extracto = uploaded_file_extracto.name.split('.')[-1]
    df_modificado = modificar_excel(uploaded_file_extracto, file_type_extracto)
    
    st.write("Vista previa del archivo de extracto bancario modificado:")
    st.dataframe(df_modificado)

    output_extracto = BytesIO()
    with pd.ExcelWriter(output_extracto, engine='xlsxwriter') as writer:
        df_modificado.to_excel(writer, index=False)
    output_extracto.seek(0)
    
    st.download_button(
        label="Descargar archivo de extracto bancario modificado",
        data=output_extracto,
        file_name="mov_modificado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if uploaded_file_pagos is not None:
    unpaid_pagos, paid_pagos = procesar_pagos(uploaded_file_pagos)
    
    st.write("Vista previa del archivo de pagos no pagados:")
    st.dataframe(unpaid_pagos)
    
    st.write("Vista previa del archivo de pagos pagados:")
    st.dataframe(paid_pagos)
    
    output_unpaid_pagos = BytesIO()
    with pd.ExcelWriter(output_unpaid_pagos, engine='xlsxwriter') as writer:
        unpaid_pagos.to_excel(writer, index=False)
    output_unpaid_pagos.seek(0)
    
    output_paid_pagos = BytesIO()
    with pd.ExcelWriter(output_paid_pagos, engine='xlsxwriter') as writer:
        paid_pagos.to_excel(writer, index=False)
    output_paid_pagos.seek(0)
    
    st.download_button(
        label="Descargar archivo de pagos no pagados",
        data=output_unpaid_pagos,
        file_name="pagos_no_pagados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        label="Descargar archivo de pagos pagados",
        data=output_paid_pagos,
        file_name="pagos_pagados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if uploaded_file_clientes is not None:
    unpaid_clientes, paid_clientes = procesar_clientes(uploaded_file_clientes)
    
    st.write("Vista previa del archivo de clientes no pagados:")
    st.dataframe(unpaid_clientes)
    
    st.write("Vista previa del archivo de clientes pagados:")
    st.dataframe(paid_clientes)
    
    output_unpaid_clientes = BytesIO()
    with pd.ExcelWriter(output_unpaid_clientes, engine='xlsxwriter') as writer:
        unpaid_clientes.to_excel(writer, index=False)
    output_unpaid_clientes.seek(0)
    
    output_paid_clientes = BytesIO()
    with pd.ExcelWriter(output_paid_clientes, engine='xlsxwriter') as writer:
        paid_clientes.to_excel(writer, index=False)
    output_paid_clientes.seek(0)
    
    st.download_button(
        label="Descargar archivo de clientes no pagados",
        data=output_unpaid_clientes,
        file_name="clientes_no_pagados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        label="Descargar archivo de clientes pagados",
        data=output_paid_clientes,
        file_name="clientes_pagados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
