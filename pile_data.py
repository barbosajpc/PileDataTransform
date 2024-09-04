# %%
import pandas as pd
import numpy as np
import streamlit as st
import time

def transform_pile_plant(data):
    # Formatar 'Position X', 'Position Y', e 'Position Z'
    data['Position X'] = data['Position X'].apply(lambda x: '{:.4f}'.format(x).replace('.', ','))
    data['Position Y'] = data['Position Y'].apply(lambda x: '{:.4f}'.format(x).replace('.', ','))
    data['Position Z'] = data['Position Z'].apply(lambda x: '{:.4f}'.format(x).replace('.', ','))
    
    # Criar a coluna 'ID_PILE'
    data['ID_PILE'] = data['C'].astype(str) + '-' + data['L'].astype(str) + data['P'].astype(str)
    
    # Reordenar as colunas para ter 'ID_PILE' primeiro
    cols = ['ID_PILE'] + [col for col in data.columns if col != 'ID_PILE']
    data_id = data[cols]
    
    # Remover 'h=' da coluna 'ALTURA'
    data_id['ALTURA'] = data_id['ALTURA'].str.replace('h=', '') 
    
    # Manter apenas as colunas especificadas
    cols_to_keep = ['ID_PILE', 'Position X', 'Position Y', 'ALTURA', 'C', 'L', 'P', 'Visibility1']
    cols_to_drop = [col for col in data_id.columns if col not in cols_to_keep]
    data_id_drop = data_id.drop(cols_to_drop, axis=1)
    
    # Renomear colunas
    data_renamed = data_id_drop.rename(columns={
        'Visibility1': 'TIPO_PILAR',
        'C': 'COLUNA', 
        'L': 'LINHA',
        'P': 'PILAR',
        'Position X': 'COORDENADA X',
        'Position Y': 'COORDENADA Y'
    }) 
    
    # Ordenar o DataFrame
    data_sorted = data_renamed.sort_values(by=['COLUNA', 'LINHA', 'PILAR'], ascending=[True, True, True])
    
    # Salvar o DataFrame resultante em um arquivo Excel
    out_path = r"C:\Users\netlu\OneDrive\Área de Trabalho\Nova pasta\pile plant\BLOCOS_PILE_PLANT_FINAL.xlsx"
    data_sorted.to_excel(out_path, index=False)

#path = r"C:\Users\netlu\OneDrive\Área de Trabalho\Nova pasta\pile plant\BLOCOS_PILE_PLANT_RAW.xls"
#data = pd.read_excel(path)
#transform_pile_plant(data)
welcome = """
# DATA TRANSFORMATION

 Data Transformation for Pile Plant foundations files in Photovoltaics Power Plants
                            in .xlsx or .csv format

"""


st.markdown(welcome,)

upload_file = st.file_uploader("Insert your file", type=["xlsx", "csv"])

if upload_file is not None:
    # Verificar a extensão do arquivo e ler o arquivo
    if upload_file.name.endswith('.xlsx'):
        data = pd.read_excel(upload_file)
    elif upload_file.name.endswith('.csv'):
        data = pd.read_csv(upload_file)

    # Exibir as primeiras linhas do DataFrame para verificação
    st.write(data.head())

st.subheader("Data pre-visualization")
st.dataframe(data.head())

st.subheader("Transformed Data pre-visualization")
st.dataframe(data.head())