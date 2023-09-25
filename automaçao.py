from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import os

# Trazer a planilha
tabela = pd.read_excel(r'C:\Users\Patrick\Desktop\programas\projetos\aut_inventario\inventario2.xlsx')

# preencher aonde está vazio com -
tabela = tabela.fillna('-')

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

navegador.get("https://forms.gle/1DrG7n13K1DCWcQKA")
qtd_linhas = tabela['NOME'].count()
print(qtd_linhas)
i = 1

while i <= qtd_linhas:
    Nome = tabela.loc[i, "NOME"]
    print(Nome)
    Modelo = tabela.loc[i, "MODELO"]
    print(Modelo)
    Usuario = tabela.loc[i, "USUÁRIO"]
    print(Usuario)
    serial_number = tabela.loc[i, "Nº DE SÉRIE"]
    print(serial_number)
    patrimonio = tabela.loc[i, "PATRIMÔNIO"]
    print(patrimonio)
    modelo_monitor = tabela.loc[i, "MODELO MONITOR"]
    print(modelo_monitor)
    serial_monitor = tabela.loc[i, "Nº DE SÉRIE MONITOR"]
    print(serial_monitor)
    patrimonio_monitor = tabela.loc[i, "PATRIMÔNIO MONITOR"]
    print(patrimonio_monitor)

    navegador.find_element(By.XPATH,
        '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(Nome)

    navegador.find_element(By.XPATH,
        '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(Modelo)

    navegador.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(Usuario)

    navegador.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(serial_number)

    navegador.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(patrimonio)

    navegador.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(modelo_monitor
    )

    navegador.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[7]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(serial_monitor)

    navegador.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[8]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(patrimonio_monitor)

    navegador.find_element(By.XPATH,
        '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span').click()

    navegador.find_element(By.XPATH,
        '/html/body/div[1]/div[2]/div[1]/div/div[4]/a').click()
    

    
    i = i+1
