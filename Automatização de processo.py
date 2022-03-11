#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# In[4]:


get_ipython().system('pip install pyautogui')
get_ipython().system('pip install pyperclip')


# In[24]:


import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 0.5
#escrever o passo a passo em português
#Passo 1  - Entrar no sistema da empresa(no caso o drive)
pyautogui.hotkey("ctrl", "t")
pyautogui.write("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.press("enter")

time.sleep(3)

#Passo 2 - Navegar no sistema até enccontrar a base de dados
pyautogui.click(x=437, y=359, clicks = 2)
time.sleep(2)

#Passo 3 - Exportar a base de vendas
pyautogui.click(x=464, y=359)
pyautogui.click(x=1655, y=227)
pyautogui.click(x=1364, y=745)
time.sleep(5)


# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[26]:


#Passo 4 - Calcular os indicadores (faturamento e quantidade de produtos vendidos)
import pandas as pd  #está avisando que pandas é pd

tabela = pd.read_excel(r"C:\Users\kello\Downloads\Jupyter\Vendas - Dez.xlsx")
display(tabela)

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()


# ### Vamos agora enviar um e-mail pelo gmail

# In[34]:


#Passo 5 - Enviar um e-mail para a diretoria com os indicadores
pyautogui.hotkey("ctrl", "t")
pyautogui.write("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(5)

#clicar no botão escrever
pyautogui.click(x=103, y=221)

#escrever destinatário
pyautogui.write("pythonimpressionador+diretoria@gmail.com")
pyautogui.press("tab")

#escrever o assunto
pyautogui.press("tab")
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")


#escrever o corpo do e-mail
pyautogui.press("tab")
texto = f""" 
Prezados, bom dia

O faturamento de ontem foi de R$ {faturamento}
e a quantidade foi de:{quantidade}

Abs.: Kelly

"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")
time.sleep(6)


#enviar o e-mail
pyautogui.hotkey("ctrl","enter")


# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# In[28]:


time.sleep(5)
pyautogui.position()


# In[ ]:





# In[ ]:




