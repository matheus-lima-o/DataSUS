import pandas as pd

# Caminho do arquivo Excel
arquivo_path = r"C:\Users\mathe\Downloads\planilhas p1\Dados_Final.xlsx"

# Carregar o arquivo Excel
df = pd.read_excel(arquivo_path)

# Substituir todos os '-' por '0' em todo o DataFrame
df = df.replace('-', '0')

# Salvar o DataFrame modificado de volta para um arquivo Excel
df.to_excel(arquivo_path, index=False)

print("Todos os hífens foram substituídos por zeros.")
