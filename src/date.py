import pandas as pd
import os

dir_dados_combinados = r'C:\Users\mathe\Downloads\planilhas p1\Dados Combinados'
arquivo_saida_final = os.path.join(dir_dados_combinados, 'Dados_Consolidados.xlsx')

# Dicionário de meses abreviados para mapeamento
meses_abrev = {
    'Jan': '01', 'Fev': '02', 'Mar': '03', 'Abr': '04',
    'Mai': '05', 'Jun': '06', 'Jul': '07', 'Ago': '08',
    'Set': '09', 'Out': '10', 'Nov': '11', 'Dez': '12'
}

# DataFrame final para acumular todos os dados
df_final = pd.DataFrame()

# Listar todos os arquivos no diretório especificado
for arquivo in os.listdir(dir_dados_combinados):
    if arquivo.startswith("Integrado_") and arquivo.endswith(".xlsx"):
        # Extrair mês e ano do nome do arquivo
        partes = arquivo.replace('.xlsx', '').split('_')
        mes = partes[1]
        ano = partes[2]

        # Carregar o DataFrame do arquivo
        caminho_completo = os.path.join(dir_dados_combinados, arquivo)
        df = pd.read_excel(caminho_completo)
        print(caminho_completo)
        
        # Adicionar colunas de Mês e Ano
        df['Mês'] = meses_abrev[mes]
        df['Ano'] = ano

        # Concatenar ao DataFrame final
        df_final = pd.concat([df_final, df], ignore_index=True)

# Salvar o DataFrame final num novo arquivo Excel
df_final.to_excel(arquivo_saida_final, index=False)

print(f"Todos os dados foram consolidados em: {arquivo_saida_final}")
