from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
import time

# Deixar em background
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

response = 's'
while(response == 's'):
    print('Digite a URL do processo:')
    # Inicialize o driver do navegador (Chrome)
    driver = webdriver.Chrome(options=chrome_options)

    url = input()
    driver.get(url)

    # Aguarda um tempo para o conteúdo ser carregado
    driver.implicitly_wait(10)

    time.sleep(2)

    content = driver.page_source
    soup = BeautifulSoup(content, 'html.parser')

    titulo_elements = soup.find_all('div', class_='v-toolbar__title')

    for titulo_element in titulo_elements:
        titulo_text = titulo_element.get_text(strip=True)

    titulo = titulo_element.get_text()

    linhas = [linha.strip() for linha in titulo.split('\n') if linha.strip()]

    if len(linhas) >= 2:
        tipo_documento = linhas[0]
        numero_processo = linhas[1]

    wait = WebDriverWait(driver, 10)

    textarea_interested = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "textarea[aria-label='Interessado']")))
    interessado_text = textarea_interested.get_attribute('value')

    textarea_summary = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "textarea[aria-label='Resumo']")))
    resumo_text = textarea_summary.get_attribute('value')

    driver.quit()

    # Inicializa a segunda página (Chrome)
    driver = webdriver.Chrome(options=chrome_options)

    url_tramitacoes = url+'tramitacoes'
    driver.get(url_tramitacoes)

    # Aguarda um tempo para o conteúdo ser carregado
    driver.implicitly_wait(10)

    content = driver.page_source

    # Localiza o elemento <tr class="pointer"> e clica nele
    tr_element = driver.find_element(By.CLASS_NAME, 'pointer')
    tr_element.click()

    wait = WebDriverWait(driver, 10)
    textarea_despatch = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "textarea[aria-label='Despacho']")))
    despacho_text = textarea_despatch.get_attribute('value')

    driver.quit()

    partes = despacho_text.split('\n\n')
    despacho = partes[0]

    print('')
    print('Tipo de documento:', tipo_documento)
    print('Número do processo:', numero_processo)
    print('Interessado:', interessado_text)
    print('Resumo:', resumo_text)
    print('Despacho:', despacho)
    print('')

    print('Deseja continuar? (s/n)')
    response = input()
