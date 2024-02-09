# Linha 1285 encaixar scrolls
# Linha 550 criar função do scroll
# Corpo de aspirantes 572, deveria ser 679
import pandas as pd
FILE_NAME        = 'teste.ods'
SHEET_LICENCAS   = 'Licenças'

licencas = pd.read_excel(FILE_NAME, sheet_name=SHEET_LICENCAS)
licencas['Nome de Guerra'] = pd.Series(range(1, 768)).apply(lambda x: str(x))
df_final = licencas[licencas['Número Interno'].apply(lambda x: x[0]) == '1'][['Número Interno', 'Nome de Guerra','Situação']]

for index, row in df_final.iterrows():
    print(row.values)