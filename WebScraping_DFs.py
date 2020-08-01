import os
import time

import pandas as pd
import unidecode
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait as wait
from tqdm import tqdm
from webdriver_manager.chrome import ChromeDriverManager

options = webdriver.ChromeOptions()
options.add_argument('--no-sandbox')
options.add_argument('--disable-gpu')

driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
# driver = webdriver.Chrome("C:\\Users\\Owner\\PycharmProjects\\RicardoCardoso25052020\\chromedriver.exe", options=options)

# Open Chrome and set initial page information
url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/CiaAb/FormBuscaCiaAb.aspx?TipoConsult=c'
driver.get(url)

# Importing the database and making a list of the CNPJ
df = pd.read_excel("Empresas_listadas_B3-CORRETO.xlsx")

lista_cnpjs = df['CNPJ']
lista_cnpjs_non_cap = []

for cnpj in tqdm(lista_cnpjs):
    current_cnpj_in_loop = cnpj.replace(".", "").replace("/", "").replace("-", "")
    driver.get(url)
    try:
        field = driver.find_element_by_id('txtCNPJNome')
        field.clear()
        field.send_keys(current_cnpj_in_loop)

        cont = driver.find_element_by_id('btnContinuar')
        cont.click()

        nome = driver.find_element_by_id('dlCiasCdCVM__ctl1_Linkbutton1')
        empresa = unidecode.unidecode(str(nome.text))

        path = "./folder/{}".format(empresa)
        # path = 'C:\\Users\\Owner\\Desktop\\DFs\\{}'.format(folder_to_create)

        if os.path.exists(path):
            pass
        else:
            os.mkdir(path)

        nome.click()
        time.sleep(10)

        periodo = driver.find_element_by_id('rdPeriodo')
        periodo.click()

        time.sleep(5)

        data_inicial = driver.find_element_by_id('txtDataIni')
        data_inicial.send_keys('04/01/2010')

        hora_inicial = driver.find_element_by_id('txtHoraIni')
        hora_inicial.send_keys('00:00')

        data_final = driver.find_element_by_id('txtDataFim')
        data_final.send_keys('22/05/2020')

        hora_final = driver.find_element_by_id('txtHoraFim')
        hora_final.send_keys('00:00')
        hora_final.click()

        # Categoria
        categories_options = driver.find_element_by_id('cboCategoria')
        options = [i.get_attribute('text') for i in categories_options.find_elements_by_tag_name('option')]

        categ = driver.find_element_by_class_name('chosen-single')
        categ.click()
        categ = driver.find_element_by_id('cboCategoria_chosen_input')
        option = [i for i in options if i == 'DFP'][0]
        categ.send_keys(option)
        categ.send_keys(Keys.RETURN)

        # Consulta
        consulta = driver.find_element_by_id('btnConsulta')
        consulta.click()

        time.sleep(5)

        page = BeautifulSoup(driver.page_source, 'lxml')

        tableid = page.find('table', id='grdDocumentos')
        fulltable = pd.read_html(str(tableid), header=0)[0]
        ativos = fulltable[fulltable["Status"] == "Ativo"].index

        visus = driver.find_elements_by_id('VisualizarDocumento')
        visus = [visus[j] for j in ativos]

        print('----> empresa atual: ' + empresa)
        for j in tqdm(visus):
            handles = driver.window_handles
            j.click()
            wait(driver, 60).until(EC.new_window_is_opened(handles))

            # Change control to opened window
            new_window = driver.window_handles[1]
            driver.switch_to.window(new_window)

            # Choosing the DF individual
            df_ind = wait(driver, 60).until(EC.presence_of_element_located((By.ID, 'cmbGrupo')))

            # df_ind = driver.find_element_by_id('cmbGrupo')
            options = [i.get_attribute("text") for i in df_ind.find_elements_by_tag_name("option")]
            values = [i.get_attribute("value") for i in df_ind.find_elements_by_tag_name("option")]

            drop = Select(df_ind)
            drop.select_by_value(values[1])
            time.sleep(10)

            # Choosing Balanco do Ativo
            balanco_ativo = driver.find_element_by_id('cmbQuadro')
            options = [i.get_attribute("text") for i in balanco_ativo.find_elements_by_tag_name("option")]
            values = [i.get_attribute("value") for i in balanco_ativo.find_elements_by_tag_name("option")]

            drop = Select(balanco_ativo)
            drop.select_by_value(values[0])  #
            time.sleep(10)

            ## You have to switch to the iframe: ##
            wait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it("iFrameFormulariosFilho"))

            page = BeautifulSoup(driver.page_source, 'lxml')

            table = page.find('table', id='ctl00_cphPopUp_tbDados')
            balanco_atv = pd.read_html(str(table), header=0)[0]

            # ano, obtido a partir do nome da coluna do dataframe
            ano = balanco_atv.columns[2][-4:]

            path2 = path + '/' + '{}'.format(ano)
            if os.path.exists(path2):
                pass
            else:
                os.mkdir(path2)

            ## Switch back to the "default content" (out of iframe) ##
            driver.switch_to.default_content()

            # Choosing Balanco do Passivo
            balanco_passivo = driver.find_element_by_id('cmbQuadro')
            options = [i.get_attribute("text") for i in balanco_passivo.find_elements_by_tag_name("option")]
            values = [i.get_attribute("value") for i in balanco_passivo.find_elements_by_tag_name("option")]

            drop = Select(balanco_passivo)
            drop.select_by_value(values[1])
            time.sleep(10)

            ## Switching to balanço passivo iframe ##
            wait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it("iFrameFormulariosFilho"))

            page = BeautifulSoup(driver.page_source, 'lxml')

            table = page.find('table', id='ctl00_cphPopUp_tbDados')
            balanco_pass = pd.read_html(str(table), header=0)[0]

            ## Switch back to the "default content" (that is, out of the iframes) ##
            driver.switch_to.default_content()

            # Chosing the DRE
            dre = driver.find_element_by_id('cmbQuadro')
            options = [i.get_attribute("text") for i in dre.find_elements_by_tag_name("option")]
            values = [i.get_attribute("value") for i in dre.find_elements_by_tag_name("option")]

            drop = Select(dre)
            drop.select_by_value(values[2])  #
            time.sleep(10)

            ## You have to switch to the iframe of the balanço passivo: ##
            wait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it("iFrameFormulariosFilho"))

            page = BeautifulSoup(driver.page_source, 'lxml')

            table = page.find('table', id='ctl00_cphPopUp_tbDados')
            demonstra_res = pd.read_html(str(table), header=0)[0]

            dfs_list = [balanco_atv, balanco_pass, demonstra_res]
            new_df = pd.concat(dfs_list)  # CONCATENATE DATAFRAMES

            # exportando um csv - a partir de algum dataframe previamente existente - para a pasta já definida (path):
            csv_name = '/' + '{}'.format(ano) + '.csv'
            new_df.to_csv(path2 + csv_name)

            driver.close()

            new_window = driver.window_handles[0]  # CHANGE TO FIRST WINDOW
            driver.switch_to.window(new_window)
        driver.get(url)
    except:
        lista_cnpjs_non_cap.append(current_cnpj_in_loop)
        driver.get(url)
        continue
