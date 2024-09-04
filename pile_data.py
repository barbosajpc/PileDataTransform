import pandas as pd
import streamlit as st

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
    out_path = "BLOCOS_PILE_PLANT_FINAL.xlsx"  # Caminho relativo
    data_sorted.to_excel(out_path, index=False)
    
    return out_path

# Interface do Streamlit
welcome = """
# DATA TRANSFORMATION

 Data Transformation for Pile Plant foundations files in Photovoltaics Power Plants
                            in .xlsx, .xls, or .csv format

"""

st.markdown(welcome)

upload_file = st.file_uploader("Insert your file", type=["xlsx", "xls", "csv"])

if upload_file is not None:
    # Verificar a extensão do arquivo e ler o arquivo
    if upload_file.name.endswith('.xlsx'):
        data = pd.read_excel(upload_file, engine='openpyxl')
    elif upload_file.name.endswith('.xls'):
        data = pd.read_excel(upload_file, engine='xlrd')  # Usar 'xlrd' para arquivos .xls
    elif upload_file.name.endswith('.csv'):
        data = pd.read_csv(upload_file)

    # Exibir as primeiras linhas do DataFrame para verificação
    st.subheader("Data pre-visualization")
    st.dataframe(data.head())

    # Botão para transformar os dados
    if st.button("Transform Pile Plant Data"):
        # Chamar a função de transformação e obter o caminho do arquivo
        out_path = transform_pile_plant(data)

        # Ler o arquivo transformado para visualização
        transformed_data = pd.read_excel(out_path)

        st.subheader("Transformed Data Pre-Visualization")
        st.dataframe(transformed_data.head())

        # Oferecer o arquivo transformado para download
        with open(out_path, "rb") as file:
            st.download_button(
                label="Download Transformed File",
                data=file,
                file_name="BLOCOS_PILE_PLANT_FINAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
