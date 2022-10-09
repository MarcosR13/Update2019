from selenium import webdriver # permite criar o navegador
from selenium.webdriver.common.keys import Keys # permite você escrever no navegador
from selenium.webdriver.common.by import By # permite você selecionar itens do navegador


# rodar em 2° plano
# from selenium.webdriver.chrome.options import Options
# chrome_options = Options()
# chrome_options.headless = True
# navegador = webdriver.Chrome(options=chrome_options)
# abrir o navegador
navegador = webdriver.Chrome()

# acessar o site
navegador.get("https://www.google.com.br/")

# pesquisar no google por cotação dolar, sempre usar aspas simples para xpath
navegador.find_element('xpath',
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação dolar")
navegador.find_element('xpath',
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_dolar = navegador.find_element('xpath',
                       '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_dolar)

# pesquisar no google por cotação euro
navegador.get("https://www.google.com.br/")
navegador.find_element('xpath',
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação euro")
navegador.find_element('xpath',
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_euro = navegador.find_element('xpath',
                       '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao_euro)

# pesquisar no google por cotação euro
navegador.execute_script("window.open('');") #abrir nova aba
navegador.switch_to.window(navegador.window_handles[1]) #definir o n° de aba
navegador.get("https://www.melhorcambio.com/ouro-hoje")
cotacao_ouro = navegador.find_element('xpath',
                       '//*[@id="comercial"]').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(",", ".")
print(cotacao_ouro)
navegador.quit()

# importar arquivo excel
import pandas as pd

tabela = pd.read_excel("Produtos.xlsx")
print(tabela)

# atualizar a coluna de cotação
# eu quero editar a coluna Cotação, onde a Coluna Moeda = Dólar
# tabela.loc[linha, coluna] = float(cotacao_dolar)
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacao_ouro)

# atualizar a coluna de preço de compra = preço original * cotação
tabela["Preço de Compra"] = tabela["Preço de Compra"] * tabela["Cotação"]

#atualizar coluna de preço de venda = preço de compra * margem
tabela["Preço de Venda"] = tabela["Preço de Venda"] * tabela["Margem"]
print(tabela)

# exportar a base de preços atualizada
tabela.to_excel("Produtos Novo.xlsx", index=False)