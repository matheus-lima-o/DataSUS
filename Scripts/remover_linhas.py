import openpyxl
import os

# Define o diretório onde estão os arquivos Excel
diretorio = r'C:\Users\mathe\Downloads\planilhas p1\Dados Combinados'

# Lista todos os arquivos Excel no diretório
arquivos_xlsx = [arquivo for arquivo in os.listdir(diretorio) if arquivo.endswith('.xlsx')]

# Processar cada arquivo Excel
for arquivo in arquivos_xlsx:
    caminho_completo = os.path.join(diretorio, arquivo)

    try:
        # Carregar o workbook
        wb = openpyxl.load_workbook(caminho_completo)
        ws = wb.active

        # Verificar se a planilha tem mais de três linhas
        if ws.max_row > 8:
            # Remover as linhas
            ws.delete_rows(2, 7)

        # Salvar o workbook modificado
        wb.save(caminho_completo)
        print(f"As linhas do arquivo {arquivo} foram removidas.")

    except Exception as e:
        print(f"Erro ao processar o arquivo {arquivo}: {e}")

print("Processamento de todos os arquivos concluído.")
