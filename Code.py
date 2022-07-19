#!/usr/bin/env python
# coding: utf-8

# In[22]:


# criar um navegador
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from selenium.webdriver.common.by import By
import time
from pathlib import Path

nav = webdriver.Chrome()


# importar/visualizar a base de dados
tabela_produtos = pd.read_excel('./buscas.xlsx')
display(tabela_produtos)
time.PAUSE = 1


# In[23]:


def busca_google_shopping(nav, produto, termos_banidos, preco_min, preco_max):
    time.PAUSE = 1
    # entrar no google
    nav.get('https://www.google.com/')

    # tratar os valores que vieram da tabela
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_produto = produto.split(" ")

    # pesquisar o nome do produto no google
    nav.find_element(By.CSS_SELECTOR, 'body > div.L3eUgb > div.o3j99.ikrT4e.om7nvf > form > div:nth-child(1) > div.A8SBwf > div.RNNXgb > div > div.a4bIc > input').send_keys(produto, Keys.ENTER)

    # clicar na aba shopping
    elementos = nav.find_elements(By.CLASS_NAME, 'hdtb-mitem')
    for item in elementos:
        if 'Shopping' in item.text:
            item.click()
            break
    # pegar a lista de resultados da busca no google_shopping
    lista_resultados = nav.find_elements(By.CLASS_NAME, 'sh-dgr__grid-result')
    
    # para cada resultado, vai ser verificado se corresponde a todas as condiçoes
    lista_ofertas = [] # lista que a função vai me retornar como resposta
    for resultado in lista_resultados[1:]:
        nome = resultado.find_element(By.CLASS_NAME, 'Xjkr3b').text
        nome = nome.lower()

        # verificação do nome - se no nome tem algum termo banido
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra in nome:
                tem_termos_banidos = True
        
        # verificação do nome - se no nome tem todos os termos do nome do produto
        tem_todos_termos_produto = True
        for palavra in lista_termos_produto:
            if not palavra in nome:
                tem_todos_termos_produto = False

        if not tem_termos_banidos and tem_todos_termos_produto: #verificando o nome
        # se tem_termos_banidos = False e tem_todos_termos_produto = True
            try:
                preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
                preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                preco = float(preco)
                
                # verificando se o preço ta dentro do mínimo e máximo
                preco_min = float(preco_min)
                preco_max = float(preco_max)
                if preco_min <= preco <= preco_max:
                    elemento_link = resultado.find_element(By.CLASS_NAME, 'aULzUe')
                    elemento_pai = elemento_link.find_element(By.XPATH, '..')
                    link = elemento_pai.get_attribute('href')
                    lista_ofertas.append((nome, preco, link))
            except:
                continue

    return lista_ofertas


def busca_buscape(nav, produto, termos_banidos, preco_min, preco_max):
    # tratar os valores da função
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_produto = produto.split(" ")
    preco_min = float(preco_min)
    preco_max = float(preco_max)
    
    # entrar no buscape
    nav.get('https://www.buscape.com.br/')
    
    # pesquisar pelo produto no buscapé
    nav.find_element(By.CLASS_NAME, 'AutoCompleteStyle_textBox__eLv3V').send_keys(produto, Keys.ENTER)
    
    # pegar a lista de resultados da busca do buscapé
    time.sleep(5)
    lista_resultados = nav.find_elements(By.CLASS_NAME, 'Cell_Content__fT5st')
    
    # para cada resultado
    lista_ofertas = []
    for resultado in lista_resultados:
        try:
            nome = resultado.get_attribute('title')
            nome = nome.lower()
            preco = resultado.find_element(By.CLASS_NAME, 'CellPrice_MainValue__JXsj_').text
            link = resultado.get_attribute('href')
            
            # ver se ele tem algum termo banido
            tem_termos_banidos = False
            for palavra in lista_termos_banidos:
                if palavra in nome:
                    tem_termos_banidos = True
            
            # ver se ele tem todos os termos do produto
            tem_todos_termos_produto = True
            for palavra in lista_termos_produto:
                if not palavra in nome:
                    tem_todos_termos_produto = False
                    
            if not tem_termos_banidos and tem_todos_termos_produto:
                preco = preco.replace(' ', '').replace('R$', '').replace('.', '').replace(',', '.')
                preco = float(preco)
                if preco_min <= preco <= preco_max:
                    lista_ofertas.append((nome, preco, link))
        except:
            pass
    return lista_ofertas
    
    # ver se eles se encontra dentro da minha faixa de preço


# Construção da Lista de Ofertas Encontras (tabela_ofertas)

# In[24]:


tabela_ofertas = pd.DataFrame()
for linha in tabela_produtos.index:
    produto = tabela_produtos.loc[linha, 'Nome']
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_min = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_max = tabela_produtos.loc[linha, 'Preço máximo']
    

    lista_ofertas_google_shopping = busca_google_shopping(nav, produto, termos_banidos, preco_min, preco_max)
    if lista_ofertas_google_shopping:
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=['produto', 'preço', 'link'])
        tabela_ofertas = tabela_ofertas.append(tabela_google_shopping)
    else:
        tabela_google_shopping = None
        
    lista_ofertas_buscape = busca_buscape(nav, produto, termos_banidos, preco_min, preco_max)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['produto', 'preço', 'link'])
        tabela_ofertas = tabela_ofertas.append(tabela_buscape)
    else:
        tabela_buscape = None
nav.quit()
display(tabela_ofertas)


# Exportar a base de dados para Excel

# In[25]:


# exportar para o excel
tabela_ofertas = tabela_ofertas.reset_index(drop=True)
tabela_ofertas.to_excel('Tabela.xlsx', index=False)


# Enviando o e-mail

# In[26]:


display(tabela_ofertas.to_html)


# In[27]:


# enviar por e-mail o resultado da tabela
import win32com.client as win32

# verificando se existe alguma oferta dentro da tabela de ofertas, ou seja, se encontrei alguma oferta na minha busca
if len(tabela_ofertas.index) > 0:
    
    # vou enviar o e-mail
    outlook = win32.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)
    mail.To = 'eltoncordeirodias@gmail.com'
    mail.Subject = 'Produto(s) Encontrado(s) na faixa de preço desejada'
    mail.HTMLBody = f'''
    <p>Prezados,</p>
    <p>Encontramos alguns produtos dentro da faixa de preço desejada. Segue tabela com detalhes.</p>
    {tabela_ofertas.to_html()}
    <p>Para quaisquer dúvidas, estou à disposição.</p>
    <p>Att.,</p>
    <p>Elton</p>
    '''

    #Anexos (pode se colocar quantos quiser):
    attachment = Path.cwd() / 'Tabela.xlsx'
    mail.Attachments.Add(str(attachment))
    

    mail.Send()

