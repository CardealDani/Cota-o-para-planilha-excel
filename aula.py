from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import pandas as pd

caminho = r'chromedriver.exe'
s = Service(caminho)
navegador = webdriver.Chrome(service=s)
navegador.get('https://www.google.com/')
# xpath ---- identificador dos elementos de um site

navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/textarea').send_keys('cotação dolar',Keys.ENTER)
dolar = navegador.find_element \
    ('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

navegador.get('https://www.google.com/')
navegador.find_element \
    ('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/textarea').send_keys('cotação euro',
                                                                                               Keys.ENTER)
euro = navegador.find_element \
    ('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

navegador.get("https://www.melhorcambio.com/ouro-hoje")

ouro = navegador.find_element \
    ('xpath', '//*[@id="comercial"]').get_attribute('value')

ouro = ouro.replace(',','.')



tabela = pd.read_excel('Produtos.xlsx')
print(tabela)


tabela.loc[tabela['Moeda'] == 'Dólar','Cotação'] = float(dolar)
tabela.loc[tabela['Moeda'] == 'Euro','Cotação'] = float(euro)
tabela.loc[tabela['Moeda'] == 'Ouro','Cotação'] = float(ouro)

tabela['Preço de Compra'] = tabela['Cotação'] * tabela['Preço Original']
tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem']
print(tabela)

tabela.to_excel('Produtos_Atualizados.xlsx',index = False)



