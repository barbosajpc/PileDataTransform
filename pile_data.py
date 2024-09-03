# %%
import pandas as pd
import numpy as np
# %%
path = r"C:\Users\netlu\OneDrive\ENERGIA SOLAR\PROJETOS\SFCR\PROJETOS 2024\UFV JUNCO I E II\PROJETOS EXECUTIVOS\PLANILHAS EXPERIMENTAIS\testes blocos dinamicos\BLOCOS_PILE_PLANT_RAW.xls"

data = pd.read_excel(path)

# %%
data['ID_PILE'] = data['C'].astype(str) + '-' + data['L'].astype(str) + data['P'].astype(str)

# %%
cols = ['ID_PILE'] + [col for col in data.columns if col != 'ID_PILE']
data_id = data[cols]
data_id.head()
# %%
cols_to_keep = 
data_id_drop = data_id.drop('Name', axis=1)
# %%
data_renamed = data_id_drop.rename(columns={'Visibility1':'TIPO_PILAR',
'C':'COLUNAS', 
'L':'LINHAS',
'P':'PILAR'}) 
# %%
data_sorted = data_renamed.sort_values(by = ['COLUNAS','LINHAS','PILAR'], ascending = [True, True, True])
