from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.hyperlink import Hyperlink
from bs4 import BeautifulSoup
import time

def initialize_driver():
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    return webdriver.Chrome(options=chrome_options)

def get_page_content(driver, url):
    driver.get(url)
    driver.implicitly_wait(10)
    time.sleep(2)
    return driver.page_source

def extract_process_info(content):
    soup = BeautifulSoup(content, 'html.parser')
    titulo_elements = soup.find_all('div', class_='v-toolbar__title')
    for titulo_element in titulo_elements:
        titulo_text = titulo_element.get_text(strip=True)
    titulo = titulo_element.get_text()
    linhas = [linha.strip() for linha in titulo.split('\n') if linha.strip()]
    if len(linhas) >= 2:
        tipo_documento = linhas[0]
        numero_processo = linhas[1]
        return tipo_documento, numero_processo
    return None, None

def extract_textarea_value(driver, css_selector):
    textarea_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, css_selector)))
    return textarea_element.get_attribute('value')

def main():
    response = 's'
    while response == 's':
        print('Digite a URL do processo:')
        url = input()

        # Abre a página do processo em background
        driver = initialize_driver()
        content = get_page_content(driver, url)
        tipo_documento, numero_processo = extract_process_info(content)

        interessado_text = extract_textarea_value(driver, "textarea[aria-label='Interessado']")
        resumo_text = extract_textarea_value(driver, "textarea[aria-label='Resumo']")

        # Fecha a página
        driver.quit()

        # Abre a página e tramitações do processo
        driver = initialize_driver()
        url_tramitacoes = url + 'tramitacoes'
        content = get_page_content(driver, url_tramitacoes)

        # Clina no elemento da página para que o conteúdo do processo apareça
        tr_element = driver.find_element(By.CLASS_NAME, 'pointer')
        tr_element.click()

        despacho_text = extract_textarea_value(driver, "textarea[aria-label='Despacho']")

        # Fecha a página
        driver.quit()

        despacho = despacho_text.split('\n\n')[0]

        dados = [
            {
                "Tipo de documento": tipo_documento,
                "Número do processo": numero_processo,
                "Interessado": interessado_text,
                "Resumo": resumo_text,
                "Despacho": despacho
            }
        ]

        file_name = "processos-lepisma.xlsx"

        try:
            workbook = load_workbook(file_name)
            sheet = workbook.active
        except FileNotFoundError:
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Tipo de documento", "Número do processo", "Interessado", "Resumo", "Despacho"])

        for dado in dados:
            hyperlink = Hyperlink(ref=f"A{sheet.max_row + 1}", target=url)
            sheet.cell(row=sheet.max_row + 1, column=1, value=dado["Tipo de documento"]).hyperlink = hyperlink
            sheet.cell(row=sheet.max_row, column=2, value=dado["Número do processo"])
            sheet.cell(row=sheet.max_row, column=3, value=dado["Interessado"])
            sheet.cell(row=sheet.max_row, column=4, value=dado["Resumo"])
            sheet.cell(row=sheet.max_row, column=5, value=dado["Despacho"])

        workbook.save(filename=file_name)

        print('Dados processados.')
        print('Deseja continuar? (s/n)')
        response = input()

if __name__ == "__main__":
    main()