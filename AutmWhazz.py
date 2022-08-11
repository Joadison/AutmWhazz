import pandas as pd
import pyautogui as pg
import os
import time
import re
import datetime, random

contatos = pd.read_excel('OK.xlsx')

time.sleep(2)

for i, telefone in enumerate(contatos['Telefone de contato']):
    pessoas = contatos.loc[i, "Nome afetado"]
    chamado = contatos.loc[i, "Referência"]
    item = contatos.loc[i, "Nome do Item"]
    numero = telefone
    numero = re.sub("\+|\(|\)|\.|\-","", str(numero))
    numero = numero.replace(" ","")
    #numero = numero [:-1]
    telefone = '55'+numero
    mensagem = '\n'
    hora = datetime.datetime.now().hour
    if hora < 12:
        mensagem = 'Bom dia, '
    elif 12 <= hora < 18:
        mensagem = 'Boa tarde, '
    else:
         mesnagem = 'Boa noite, '
    print(telefone)
    print(chamado)
    print(pessoas)
    mensagem += '''%s! Meu nome é Joadison! Sou do CATI, estou entrando em contato referente ao chamado.:%s *%s*'''%(pessoas, chamado, item)
    texto = mensagem.replace(" ", "%20")
    os.system('start whatsapp://send?phone=%s"&"text=%s' %(telefone, texto))
    time.sleep(10)
    pg.click()
    pg.press('Enter')
    time.sleep(10)

    #os.system('start chrome "https://wa.me/send?phone=%s&text=%s"' %(telefone, mensagem))
