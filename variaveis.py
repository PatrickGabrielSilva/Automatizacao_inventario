import pandas as pd
import pyautogui

tabela = pd.read_excel("C:/Users/Patrick/Desktop/projetos/aut_inventario/inventario2.xlsx")


for i, Nome in enumerate(tabela["NOME"]):
    Modelo = tabela.loc[i, "MODELO"]
    Usuario = tabela.loc[i,"USUÁRIO"]
    serial_number = tabela.loc[i,"Nº DE SÉRIE"]
    patrimonio = tabela.loc[i,"PATRIMÔNIO"]
    modelo_monitor = tabela.loc[i,"MODELO MONITOR"]
    serial_monitor = tabela.loc[i,"Nº DE SÉRIE MONITOR"]
    patrimonio_monitor = tabela.loc[i,"PATRIMÔNIO MONITOR"]