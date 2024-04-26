from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Configurações do navegador
options = Options()
options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Users\mathe\Downloads\planilhas p1\Quantidade Aprovada",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

# Configura o WebDriver
driver = webdriver.Chrome(options=options)
driver.get("http://tabnet.datasus.gov.br/cgi/deftohtm.exe?sih/cnv/spabr.def")

# Espera a página carregar
wait = WebDriverWait(driver, 10)

# Seleciona 'Grupo procedimento' no primeiro dropdown
dropdown = Select(wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'div.coluna select'))))
dropdown.select_by_visible_text('Subgrupo proced.')

# 'Loop' pelos anos e meses
for ano in range(19, 24):  # De 2019 a 2023
    for mes in range(1, 13):  # De Janeiro a Dezembro

        # Constrói o código do arquivo para o mês e ano corrente
        codigo = f"spabr{ano}{mes:02d}.dbf"

        # Print para conferir o andamento do download dos arquivos
        print(codigo)

        # Seleciona o arquivo no segundo dropdown
        dropdown_arquivos = Select(wait.until(EC.element_to_be_clickable((By.NAME, 'Arquivos'))))
        dropdown_arquivos.deselect_all()
        dropdown_arquivos.select_by_value(codigo)

        # Clica no botão 'Mostra'
        mostra_button = wait.until(EC.element_to_be_clickable((By.NAME, 'mostre')))
        mostra_button.click()

        # Muda para a nova aba que foi aberta
        driver.switch_to.window(driver.window_handles[-1])

        # Localiza e clica no botão 'Copia como .CSV'
        try:
            csv_link = wait.until(
                EC.element_to_be_clickable((
                    By.XPATH, "//td[@class='botao_opcao']/a[contains(text(), 'Copia como .CSV')]")))
            csv_link.click()
            # Aguarda um tempo para o download concluir
            time.sleep(2)
            # Fecha a aba atual
            driver.close()
            # Retorna para a aba original
            driver.switch_to.window(driver.window_handles[0])
        except Exception as e:
            print(f"Erro ao tentar baixar o arquivo")
            driver.close()

        if ano == 23 and mes > 12:  # Para parar em Dez/2023
            break

# Configura uma espera longa para não fechar o navegador imediatamente (para testes)
# time.sleep(100000)

# Fecha o navegador
driver.quit()

