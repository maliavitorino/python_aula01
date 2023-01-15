#!/usr/bin/env python
# coding: utf-8

# In[2]:


get_ipython().system('pip install pyautogui')
import pyautogui
import pyperclip # python tem a limitação de caracteres especiais 
import time # para controlar o tempo de pausa entre os comandos

# Passo 1: Entrar no sistema da empresa (no link do drive)
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://drive.google.com/drive/u/0/my-drive") # import pyperclip 
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(10) # import time

# Passo 2: Navegar até o Local do relatório (entrar na pasta Exportar)
pyautogui.click(x=359, y=498)
time.sleep(10)

# Passo 3: Exportar o relatório (fazer o download)
pyautogui.click(x=1177, y=186) # três pontinhos
pyautogui.click(1100, y=551) # download
time.sleep(10)

# Passo 4: Calcular os indicadores (fazer o download) 
import pandas as pd # um dos pacotes mais usado com milhares de infos eficientes; trazer bases de dados e analisar no python
import os # pesquisar sobre

tabela = pd.read_excel(r"C:\Users\Vande\Downloads\Vendas - Dez.xlsx") 
# colocar sempre um r de raw antes de um caminho de arquivo por causa dos caracteres especiais

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
print(faturamento)
print(quantidade)

time.sleep(10)

# Passo 5: Enviar um e-mail para a diretoria
# abrir aba e entrar no email

pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/0/?tab=rm&ogbl") # import pyperclip 
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")

time.sleep(10)

# clicar no o botão escrever

pyautogui.click(x=67, y=205) # botão escrever

time.sleep(10)

# preencher as infos

pyperclip.copy("ahrimalia@gmail.com")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter") # seleciona o email
pyautogui.press("tab") # pula p o campo assunto

# assunto
pyperclip.copy("Teste Python")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
                
#corpo

texto = f"""
Prezada Malia, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: R${quantidade:,}

Att,
Malia""" # aspas triplas para selecionar texto longo e F antes das 3 aspas para poder puxar as variáveis no texto # {} para puxar a variável

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# enviar o email
pyautogui.hotkey("ctrl","enter")

