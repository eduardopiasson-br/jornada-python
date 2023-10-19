# ------------------------------------------------ PASSO A PASSO DO PROJETO ------------------------------------------------ #
# pip install pyautogui pandas numpy openpyxl
import pyautogui
import time

# ----------------------------------------- Passo 1 : Entrar no sistema da empresa ----------------------------------------- #
# https://dlp.hashtagtreinamentos.com/python/intensivao/login
# pyautogui.click -> clica com mouse
# pyautogui.write -> escreve texto
# pyautogui.press -> aperta 1 tecla
# pyautogui.hotkey -> atalho (combina teclas)

pyautogui.PAUSE = 0.3

# abrir chrome
pyautogui.press("win")
pyautogui.write("Chrome")
pyautogui.press("enter")

# entrar no link
link = "https://dlp.hashtagtreinamentos.com/python/intensivao/login"
pyautogui.write(link)
pyautogui.press("enter")

# esperar site carregar
time.sleep(3)

# ------------------------------------------------- Passo 2 : Fazer Login ------------------------------------------------- #
pyautogui.click(x=824, y=453) # Email
pyautogui.write("eduardo@gmail.com") # Emai

pyautogui.press("tab") # Senha
pyautogui.write("123123") # Senha

pyautogui.press("tab") # Button
pyautogui.press("enter") # Button

# esperar login efetuar
time.sleep(3)

# ------------------------------------- Passo 3 : Importar base de dados de produtos -------------------------------------- #
import pandas 

tabela = pandas.read_csv('produtos.csv')
# print(tabela)

# ------------------------------------------- Passo 4 : Cadastrar um produto ---------------------------------------------- #

pyautogui.click(x=761, y=313)
pyautogui.write("MOLO000251") 
pyautogui.press("tab")
pyautogui.write("Logitech")
pyautogui.press("tab")
pyautogui.write("Mouse")
pyautogui.press("tab")
pyautogui.write("1")
pyautogui.press("tab")
pyautogui.write("25.95")
pyautogui.press("tab")
pyautogui.write("6.50")
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.press("enter")

pyautogui.scroll(50000)

# --------------------------------- Passo 5 : Repetir o cadastro para todos os produtos ----------------------------------- #

for linha in tabela.index:
    pyautogui.click(x=761, y=313)

    pyautogui.write(str(tabela.loc[linha, "codigo"])) 
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "marca"]))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "tipo"]))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "categoria"]))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "preco_unitario"]))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "custo"]))
    pyautogui.press("tab")
    obs = tabela.loc[linha, "obs"]
    if not pandas.isna(obs):        
        pyautogui.write(str(obs))
    pyautogui.press("tab")
    pyautogui.press("enter")

    pyautogui.scroll(50000)

print("Automação Executada com Sucesso!")