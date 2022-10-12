import pandas as pd
import os

#Ler arquivos do caminho

path = r"Passar a pasta com arquivos"
files = os.listdir(path)
df = pd.DataFrame()

files_xlsx = (path + '\\' + f for f in files if f[-4:0] == 'formado do arquivo')

for f in files_xlsx:
    data = pd.read_excel(f)
    df = df.append(data)

df.to_excel(r'Passar caminho de saída')

