#!/usr/bin/env python
# coding: utf-8

# In[45]:


from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By


options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Users\Thomas\downloads",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=options)


# In[46]:


import os

caminho = os.getcwd()
arquivo = caminho + r'\login.html'
navegador.get(arquivo)


# In[47]:


#Entrar na página de Login
#Senha e Usuário
navegador.find_element(By.XPATH, '/html/body/div/form/input[1]').send_keys('thomas@gmail.com')
navegador.find_element(By.XPATH, '/html/body/div/form/input[2]').send_keys('123456')
#Clicar no botão Login
navegador.find_element(By.XPATH, '/html/body/div/form/button').click()


# In[48]:


#Importar Base de Dados

import pandas as pd

clientes_df = pd.read_excel('NotasEmitir.xlsx')


display(clientes_df)


# In[49]:



#Percorrer as linhas do dataframe

for linha in clientes_df.index:
    
    #Preencher Dados


    #Nome Razão Social
    navegador.find_element(By.NAME, 'nome').send_keys(clientes_df.loc[linha, 'Cliente'])

    #Endereço
    navegador.find_element(By.NAME, 'endereco').send_keys(clientes_df.loc[linha, 'Endereço'])

    #Bairro
    navegador.find_element(By.NAME, 'bairro').send_keys(clientes_df.loc[linha, 'Bairro'])

    #Municipio
    navegador.find_element(By.NAME, 'municipio').send_keys(clientes_df.loc[linha, 'Municipio'])

    #cep
    navegador.find_element(By.NAME, 'cep').send_keys(str(clientes_df.loc[linha, 'CEP']))
    
    #uf
    navegador.find_element(By.NAME, 'uf').send_keys(clientes_df.loc[linha, 'UF'])
    #Cnpj

    navegador.find_element(By.NAME, 'cnpj').send_keys(str(clientes_df.loc[linha, 'CPF/CNPJ']))

    #Inscrição Estadual

    navegador.find_element(By.NAME, 'inscricao').send_keys(str(clientes_df.loc[linha, 'Inscricao Estadual']))

    #Descrição

    navegador.find_element(By.NAME, 'descricao').send_keys(clientes_df.loc[linha, 'Descrição'])


    #quantidade

    navegador.find_element(By.NAME, 'quantidade').send_keys(str(clientes_df.loc[linha, 'Quantidade']))

    #valor unitario

    navegador.find_element(By.NAME, 'valor_unitario').send_keys(str(clientes_df.loc[linha, 'Valor Unitario']))

    #valor total

    navegador.find_element(By.NAME, 'total').send_keys(str(clientes_df.loc[linha, 'Valor Total']))

    #Clicar Emitir Nota

    navegador.find_element(By.CLASS_NAME, 'registerbtn').click()
    
    #recarregar a página
    
    navegador.refresh()


# In[ ]:


navegador.quit()


# In[ ]:





# In[ ]:





# In[ ]:




