#!/usr/bin/env python
# coding: utf-8

# In[175]:


from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import pandas as pd
import re


# In[176]:


#Importing the database and making a list of the CNPJ
df = pd.read_excel("Extracao_Divida_2020.xlsx")
df.head()
# print(df['CNPJ'].tolist())
doc = df['CNPJ'][1]
doc


# In[177]:


options = webdriver.ChromeOptions()
options.add_argument('--no-sandbox')
options.add_argument('--disable-gpu')

driver = webdriver.Chrome(options=options)


# In[178]:


url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/CiaAb/FormBuscaCiaAb.aspx?TipoConsult=c'
driver.get(url)


# In[179]:


field = driver.find_element_by_id('txtCNPJNome')
field.clear()
field.send_keys('61.079.117/0001-05')


# In[180]:


cont = driver.find_element_by_id('btnContinuar')
cont.click()


# In[181]:


#if driver.find_element_by_id('lblMsg'):
#    df['status'] = 'nao listada'
#    print('encontrou')


# In[182]:


nome = driver.find_element_by_id('dlCiasCdCVM__ctl1_Linkbutton1')
nome.click()


# In[183]:


periodo = driver.find_element_by_id('rdPeriodo')
periodo.click()


# In[184]:


data_inicial = driver.find_element_by_id('txtDataIni')
data_inicial.send_keys('04/01/2010')

hora_inicial = driver.find_element_by_id('txtHoraIni')
hora_inicial.send_keys('00:00')

data_final = driver.find_element_by_id('txtDataFim')
data_final.send_keys('22/05/2020')

hora_final = driver.find_element_by_id('txtHoraFim')
hora_final.send_keys('00:00')


# In[185]:


#Categoria
categories_options = driver.find_element_by_id('cboCategoria')
options = [i.get_attribute('text') for i in categories_options.find_elements_by_tag_name('option')]

categ = driver.find_element_by_class_name('chosen-single')
categ.click()
categ = driver.find_element_by_id('cboCategoria_chosen_input')
categ.send_keys(options[10])
categ.send_keys(Keys.RETURN)


# In[186]:


#Consulta
consulta = driver.find_element_by_id('btnConsulta')
consulta.click()


# In[187]:


#visualizar documento
visu = driver.find_element_by_id('VisualizarDocumento')
visu.click()


# In[ ]:




