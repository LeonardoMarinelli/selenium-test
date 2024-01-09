from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pyautogui as pause
from selenium.webdriver.common.by import By
from datetime import date
from pathlib import Path
from openpyxl import Workbook
import os

def getDolarAndEuro():
    option = webdriver.ChromeOptions()
    option.add_argument('headless')

    navegador = webdriver.Chrome(options=option)

    navegador.get('https://economia.uol.com.br/cotacoes/cambio/')
    pause.sleep(2)

    valorDolar = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div/section[1]/div[1]/a/div/span[2]').text
    valorEuro = navegador.find_element(By.XPATH, '/html/body/div[1]/div/div/section[1]/div[3]/a/div/span[2]').text

    valorDolar = valorDolar.replace(',', '.')
    valorEuro = valorEuro.replace(',', '.') 

    valorDolar = round(float(valorDolar), 2)
    valorEuro = round(float(valorEuro), 2)

    navegador.quit()

    return valorDolar, valorEuro

def writeXLSX(dolar, euro):
    workbook = Workbook()
    worksheet = workbook.active

    worksheet['A1'] = 'VALOR DOLAR'
    worksheet['B1'] = 'VALOR EURO'
    worksheet['C1'] = 'DATA'
    worksheet['A2'] = dolar
    worksheet['B2'] = euro
    worksheet['C2'] = date.today().strftime('%d/%m/%Y')

    for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

    script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
    file_path = script_dir / 'cotacao.xlsx'
    workbook.save(file_path)

    workbook.save(file_path)
    print(f'Arquivo salvo em: {file_path}')

dolar, euro = getDolarAndEuro()
dolar, euro = str(dolar).replace('.', ','), str(euro).replace('.', ',')
writeXLSX(dolar, euro)