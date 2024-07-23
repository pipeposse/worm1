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

st.title("Modificador de Archivos Excel")

uploaded_file_extracto = st.file_uploader("Subí el último extracto bancario", type=["xls", "xlsx"])
uploaded_file_pagos = st.file_uploader("Subí el archivo de pagos", type=["xls", "xlsx"])

if uploaded_file_extracto is not None and uploaded_file_pagos is not None:
    file_type_extracto = uploaded_file_extracto.name.split('.')[-1]
    df_modificado = modificar_excel(uploaded_file_extracto, file_type_extracto)
    
    st.write("Vista previa del archivo de extracto bancario modificado:")
    st.dataframe(df_modificado)

    unpaid, paid = procesar_pagos(uploaded_file_pagos)
    
    st.write("Vista previa del archivo de pagos no pagados:")
    st.dataframe(unpaid)
    
    st.write("Vista previa del archivo de pagos pagados:")
    st.dataframe(paid)

    output_extracto = BytesIO()
    with pd.ExcelWriter(output_extracto, engine='xlsxwriter') as writer:
        df_modificado.to_excel(writer, index=False)
    output_extracto.seek(0)
    
    output_unpaid = BytesIO()
    with pd.ExcelWriter(output_unpaid, engine='xlsxwriter') as writer:
        unpaid.to_excel(writer, index=False)
    output_unpaid.seek(0)
    
    output_paid = BytesIO()
    with pd.ExcelWriter(output_paid, engine='xlsxwriter') as writer:
        paid.to_excel(writer, index=False)
    output_paid.seek(0)
    
    st.download_button(
        label="Descargar archivo de extracto bancario modificado",
        data=output_extracto,
        file_name="mov_modificado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        label="Descargar archivo de pagos no pagados",
        data=output_unpaid,
        file_name="pagos_no_pagados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        label="Descargar archivo de pagos pagados",
        data=output_paid,
        file_name="pagos_pagados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
