import os

diretorio = r"C:\Users\mathe\Downloads\planilhas p1\Quantidade Aprovada"

# Mapeamento de número do mês
meses = {
    1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr',
    5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Ago',
    9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'
}

# Lista de todos os arquivos CSV no diretório
arquivos = [arq for arq in os.listdir(diretorio) if arq.endswith('.xlsx')]
arquivos.sort()

# Verificar se o número de arquivos
if len(arquivos) != 5 * 12:  # 5 anos, 12 meses, 60 arquivos
    print("Erro: O número de arquivos não corresponde ao intervalo de 2019 a 2023")

# Iniciar o mês e o ano
ano = 2019
mes = 1

# Loop pelos arquivos para renomeá-los
for i, arquivo in enumerate(arquivos):
    # Construir o novo nome do arquivo
    nome_novo = f"Quant_Aprov_Subgp_Mun_{meses[mes]}_{ano}_new.xlsx"
    nome_antigo_completo = os.path.join(diretorio, arquivo)
    nome_novo_completo = os.path.join(diretorio, nome_novo)

    # Renomear o arquivo
    os.rename(nome_antigo_completo, nome_novo_completo)

    # Atualizar mês e ano
    mes += 1
    if mes > 12:
        mes = 1
        ano += 1

print("Renomeados.")
