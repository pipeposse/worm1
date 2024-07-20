import streamlit as st
import pandas as pd
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

st.title("Modificador de Archivos Excel")

uploaded_file = st.file_uploader("Elige un archivo Excel", type=["xls", "xlsx"])

if uploaded_file is not None:
    file_type = uploaded_file.name.split('.')[-1]
    df_modificado = modificar_excel(uploaded_file, file_type)
    
    st.write("Vista previa del archivo modificado:")
    st.dataframe(df_modificado)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_modificado.to_excel(writer, index=False)
    output.seek(0)
    
    st.download_button(
        label="Descargar archivo modificado",
        data=output,
        file_name="mov_modificado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
