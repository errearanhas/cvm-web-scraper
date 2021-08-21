from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver
import pandas as pd
import numpy as np
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import time
import os
from tqdm import tqdm
from unidecode import unidecode
from webdriver_manager.chrome import ChromeDriverManager

import html5lib

options = webdriver.ChromeOptions()
options.add_argument('--no-sandbox')
options.add_argument('--disable-gpu')

driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
# driver = webdriver.Chrome("C:\\Users\\Owner\\PycharmProjects\\RicardoCardoso25052020\\chromedriver.exe",
#                           options=options)

time.sleep(10)

# Open Chrome and set initial page information

url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/CiaAb/FormBuscaCiaAb.aspx?TipoConsult=c'
driver.get(url)

# Importing the database and making a list of the CNPJ

df = pd.read_excel("Empresas_listadas_B3-CORRETO.xlsx")
df.head()
lista_cnpjs = df['CNPJ']
lista_cnpjs_non_cap = []
empresa = []
nome1 = []
email = []

# cnpj = '42.771.949/0001-35'

for cnpj in tqdm(lista_cnpjs):
    current_cnpj_in_loop = cnpj.replace(".", "").replace("/", "").replace("-", "")
    try:
        field = driver.find_element_by_id('txtCNPJNome')
        field.clear()
        field.send_keys(current_cnpj_in_loop)

        cont = driver.find_element_by_id('btnContinuar')
        cont.click()

        timer = np.random.randint(10, 20)
        time.sleep(timer)

        try:
            nome = WebDriverWait(driver, timer).until(
                EC.presence_of_element_located((By.ID, 'dlCiasCdCVM__ctl1_Linkbutton1'))
            )
        except:
            lista_cnpjs_non_cap.append(current_cnpj_in_loop)
            driver.get(url)
            continue

        nome = driver.find_element_by_id('dlCiasCdCVM__ctl1_Linkbutton1')
        nome.click()

        try:
            wait_for_btnConsulta = WebDriverWait(driver, timer).until(
                EC.presence_of_element_located((By.ID, 'btnConsulta'))
            )
        except:
            lista_cnpjs_non_cap.append(current_cnpj_in_loop)
            driver.get(url)
            continue

        time.sleep(10)

        periodo = driver.find_element_by_id('rdPeriodo')
        periodo.click()

        data_inicial = driver.find_element_by_id('txtDataIni')
        data_inicial.send_keys('04/01/2010')
        periodo.click()

        hora_inicial = driver.find_element_by_id('txtHoraIni')
        hora_inicial.send_keys('00:00')
        periodo.click()

        data_final = driver.find_element_by_id('txtDataFim')
        data_final.send_keys('14/04/2021')

        hora_final = driver.find_element_by_id('txtHoraFim')
        hora_final.send_keys('00:00')
        periodo.click()

        # get and set category options
        categ_options = driver.find_element_by_id('cboCategorias')
        options = [i.get_attribute('text') for i in categ_options.find_elements_by_tag_name('option')]

        categ = driver.find_element_by_class_name('chosen-single')
        categ.click()
        categ = driver.find_element_by_class_name('chosen-search-input')
        option = [i for i in options if 'Formulário Cadastral' in i][0]

        categ.send_keys(option)
        categ.send_keys(Keys.RETURN)

        # Consulta
        consulta = driver.find_element_by_id('btnConsulta')
        consulta.click()
        time.sleep(10)

        page = BeautifulSoup(driver.page_source, 'lxml')

        # visus = driver.find_elements_by_id('VisualizarDocumento')
        j = driver.find_element_by_id('VisualizarDocumento')
        j.click()
        time.sleep(10)

        # Change to control the second window that poped

        new_window = driver.window_handles[-1]  # CHANGE TO SECOND WINDOW
        driver.switch_to.window(new_window)

        dados_gerais = driver.find_element_by_id('cmbQuadro')
        options = [i.get_attribute("text") for i in dados_gerais.find_elements_by_tag_name("option")]
        values = [i.get_attribute("value") for i in dados_gerais.find_elements_by_tag_name("option")]

        drop = Select(dados_gerais)
        drop.select_by_value(values[-2])  # value of DRI option
        time.sleep(5)

        page = BeautifulSoup(driver.page_source, 'lxml')

        # switching to the iframe:
        driver.switch_to.frame(driver.find_element_by_tag_name("iframe"))
        greenButton = driver.find_element_by_id('ctl00_cphPopUp_imgSetaFiltro1')
        greenButton.click()

        page = BeautifulSoup(driver.page_source, 'lxml')
        droppedTable = page.find('table', id='ctl00_cphPopUp_Espacamento1')
        table = pd.read_html(str(droppedTable), header=0)[0]

        nm = unidecode(table.iloc[0, 1])
        mail = table.iloc[-1, 1]

        nome1.append(nm)
        email.append(mail)

        driver.switch_to.default_content()

        new_window = driver.window_handles[0]  # CHANGE TO PREVIOUS WINDOW

        driver.close()
        driver.switch_to.window(new_window)
        driver.back()

        timer = np.random.randint(10, 20)
        time.sleep(timer)

        company = driver.find_element_by_id('dlCiasCdCVM__ctl1_Linkbutton1')
        empresa.append(unidecode(company.text))

        dtframe = pd.DataFrame({'empresa': empresa, 'nome': nome1, 'email': email})
        dtframe.to_csv('tabela_empresa_nomes.csv', index=False)

        driver.get(url)
        timer = np.random.randint(10, 20)
        time.sleep(timer)
    except:
        lista_cnpjs_non_cap.append(current_cnpj_in_loop)
        driver.get(url)
        timer = np.random.randint(10, 20)
        time.sleep(timer)
        continue
