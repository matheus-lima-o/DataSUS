import pandas as pd

caminho_DataSUS = r"C:\Users\mathe\Downloads\planilhas p1\Dados_Consolidados.xlsx"
caminho_DataCenso = r"C:\Users\mathe\Downloads\planilhas p1\CENSO_DEMOGRAFICO_2022.xlsx"

# Carregar o arquivo Excel
df_DataSUS = pd.read_excel(caminho_DataSUS)

# Carregar o arquivo CSV
df_Censo = pd.read_excel(caminho_DataCenso)

# Coluna comum
coluna_comum = 'MUNICÍPIO'

# Verificar se a coluna existe em ambos os DataFrames
if coluna_comum in df_DataSUS.columns and coluna_comum in df_Censo.columns:
    # Realizar o merge
    df_merged = pd.merge(df_DataSUS, df_Censo, on=coluna_comum, how='outer')

    # Salvar o DataFrame combinado num novo arquivo Excel
    caminho_saida = r"C:\Users\mathe\Downloads\planilhas p1\Dados_Final.xlsx"
    df_merged.to_excel(caminho_saida, index=False)

    print(f"Merge completado e arquivo salvo em: {caminho_saida}")
else:
    print(f"A coluna '{coluna_comum}' não foi encontrada em ambos os DataFrames. Verifique os nomes das colunas.")