import pandas as pd
import os
from datetime import datetime

dir_quantidade = r'C:\Users\mathe\Downloads\planilhas p1\Quantidade Aprovada'
dir_valor = r'C:\Users\mathe\Downloads\planilhas p1\Valor Aprovado'
dir_saida = r'C:\Users\mathe\Downloads\planilhas p1\Dados Combinados'

# Certificar de que o diretório de saída existe
if not os.path.exists(dir_saida):
    os.makedirs(dir_saida)

# Converter número do mês para abreviatura
meses_abrev = {
    1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr',
    5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Ago',
    9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'
}

# Definir o intervalo de datas
inicio = datetime(2019, 1, 1)
fim = datetime(2023, 12, 31)

# Processar cada mês dentro do intervalo
data_atual = inicio
while data_atual <= fim:
    mes_abrev = meses_abrev[data_atual.month]
    ano = data_atual.strftime('%Y')  # ano com quatro dígitos
    nome_arquivo = f"_{mes_abrev}_{ano}.xlsx"

    caminho_quantidade = os.path.join(dir_quantidade, f'Quant_Aprov_Subgp_Mun{nome_arquivo}')
    caminho_valor = os.path.join(dir_valor, f'Valor_Aprov_Subgp_Mun{nome_arquivo}')
    caminho_saida_final = os.path.join(dir_saida, f'Integrado{nome_arquivo}')

    if os.path.exists(caminho_quantidade) and os.path.exists(caminho_valor):
        # Carregar os DataFrames
        df_quantidade = pd.read_excel(caminho_quantidade)
        df_valor = pd.read_excel(caminho_valor)

        # Realizar o merge
        df_merged = pd.merge(df_quantidade, df_valor, on='Município', how='outer')

        # Salvar o DataFrame combinado num novo arquivo Excel
        df_merged.to_excel(caminho_saida_final, index=False)
        print(f"Merge completado para {mes_abrev}/{ano} e arquivo salvo em: {caminho_saida_final}")

    else:
        print(f"Arquivos para {mes_abrev}/{ano} não encontrados.")

    # Incrementar para o próximo mês
    data_atual = pd.to_datetime(data_atual) + pd.DateOffset(months=1)

print("Processamento de todos os arquivos concluído.")
