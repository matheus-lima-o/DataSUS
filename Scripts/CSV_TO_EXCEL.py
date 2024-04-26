import os
import pandas as pd
import win32com.client as win32


def abrir_salvar_excel(caminho_original, caminho_salvo):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(caminho_original)
    wb.SaveAs(caminho_salvo, FileFormat=51)
    wb.Close()
    excel.Application.Quit()


diretorioQA_csv = r"C:\Users\mathe\Downloads\planilhas p1\Backup\Quantidade Aprovada"
diretorioQA_xlsx = r"C:\Users\mathe\Downloads\planilhas p1\Quantidade Aprovada"

# Assegurar que o diretório de destino existe
if not os.path.exists(diretorioQA_xlsx):
    os.makedirs(diretorioQA_xlsx)

# Obter lista de arquivos CSV
arquivosQA_csv = [arquivo for arquivo in os.listdir(diretorioQA_csv) if arquivo.endswith('.csv')]

# Processar cada arquivo CSV
for arquivo in arquivosQA_csv:
    caminho_completo_csv = os.path.join(diretorioQA_csv, arquivo)
    caminho_completo_xlsx = os.path.join(diretorioQA_xlsx, arquivo.replace('.csv', '.xlsx'))

    # Abrir e salvar o arquivo via Excel para normalizar
    abrir_salvar_excel(caminho_completo_csv, caminho_completo_xlsx)

    # Ler o arquivo normalizado com Pandas
    df = pd.read_excel(caminho_completo_xlsx)
    print(df.head())

print("Processamento QA concluído.")


# Diretórios
diretorioVA_csv = r"C:\Users\mathe\Downloads\planilhas p1\Backup\Valor Aprovado"
diretorioVA_xlsx = r"C:\Users\mathe\Downloads\planilhas p1\Valor Aprovado"

# Assegurar que o diretório de destino existe
if not os.path.exists(diretorioVA_xlsx):
    os.makedirs(diretorioVA_xlsx)

# Obter lista de arquivos CSV
arquivosVA_csv = [arquivo for arquivo in os.listdir(diretorioVA_csv) if arquivo.endswith('.csv')]

# Processar cada arquivo CSV
for arquivo in arquivosVA_csv:
    caminho_completo_csv = os.path.join(diretorioVA_csv, arquivo)
    caminho_completo_xlsx = os.path.join(diretorioVA_xlsx, arquivo.replace('.csv', '.xlsx'))

    # Abrir e salvar o arquivo via Excel para normalizar
    abrir_salvar_excel(caminho_completo_csv, caminho_completo_xlsx)

    # Agora ler o arquivo normalizado com Pandas
    df = pd.read_excel(caminho_completo_xlsx)
    print(df.head())  # Exibir as primeiras linhas para verificação

print("Processamento QA concluído.")


