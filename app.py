import streamlit as st
import pandas as pd
from datetime import datetime
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
    try:
        if file_type == 'xlsx':
            df = pd.read_excel(file, engine='openpyxl')
        else:
            df = pd.read_excel(file, engine='xlrd')
        
        df = convertir_columnas(df)
        df = extraer_info_adicional(df)
        return df
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
        return None

def procesar_pagos_o_clientes(file):
    try:
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

        # Crear la columna 'Estado de Deuda/Cobranza'
        df['Estado'] = df.apply(lambda row: 'Vencida' if row['Fecha de vencimiento'] < today and row['Estado de pago'] == 'No pagadas' else ('Por pagar' if row['Estado de pago'] == 'No pagadas' else row['Estado de pago']), axis=1)

        # Dividir en pagados y otros
        df_pagados = df[df['Estado de pago'] == 'Pagado']
        df_otros = df[df['Estado de pago'] != 'Pagado']

        return df_pagados, df_otros
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
        return None, None

def mostrar_y_descargar(df_pagados, df_otros, filename_prefix):
    st.write("Vista previa de pagos/cobranzas realizados:")
    st.dataframe(df_pagados.head(4))  # Mostrar solo las primeras 4 filas

    st.write("Vista previa de pagos/cobranzas no realizados:")
    st.dataframe(df_otros.head(4))  # Mostrar solo las primeras 4 filas

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_pagados.to_excel(writer, sheet_name='realizados', index=False)
        df_otros.to_excel(writer, sheet_name='no_realizados', index=False)
    output.seek(0)

    st.download_button(
        label=f"Descargar archivo {filename_prefix} modificado",
        data=output,
        file_name=f"{filename_prefix}_mod.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.title("Modificador de Archivos Excel")

uploaded_file_extracto = st.file_uploader("Subí el último extracto bancario", type=["xls", "xlsx"])
uploaded_file_pagos = st.file_uploader("Subí facturas de Proveedores", type=["xls", "xlsx"])
uploaded_file_clientes = st.file_uploader("Subí factura de Clientes", type=["xls", "xlsx"])

if uploaded_file_extracto is not None:
    file_type_extracto = uploaded_file_extracto.name.split('.')[-1]
    df_modificado = modificar_excel(uploaded_file_extracto, file_type_extracto)
    
    if df_modificado is not None:
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
    df_pagados, df_otros = procesar_pagos_o_clientes(uploaded_file_pagos)
    
    if df_pagados is not None and df_otros is not None:
        mostrar_y_descargar(df_pagados, df_otros, "proveedores")

if uploaded_file_clientes is not None:
    df_pagados, df_otros = procesar_pagos_o_clientes(uploaded_file_clientes)
    
    if df_pagados is not None and df_otros is not None:
        mostrar_y_descargar(df_pagados, df_otros, "clientes")
