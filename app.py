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
        df = extra
