import pandas as pd

file_path = "tamanhos.xlsx"

# Carrega o arquivo e exibe as colunas reais
df = pd.read_excel(file_path)
print(df.columns)