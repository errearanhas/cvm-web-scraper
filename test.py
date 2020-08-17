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
from selenium.webdriver.support.ui import WebDriverWait
from tqdm import tqdm
from webdriver_manager.chrome import ChromeDriverManager

# import Excel database and get list of cnpjs
df = pd.read_excel("Empresas_listadas_B3-CORRETO.xlsx")
lista_cnpjs = df['CNPJ']

# prepare a list to log non captured cnpjs
lista_cnpjs_non_cap = []

# set driver options and open url
options = webdriver.ChromeOptions()
options.add_argument('--no-sandbox')
options.add_argument('--disable-gpu')

driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
# driver = webdriver.Chrome("C:\\Users\\Owner\\PycharmProjects\\RicardoCardoso25052020\\chromedriver.exe", options=options)

url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/CiaAb/FormBuscaCiaAb.aspx?TipoConsult=c'
driver.get(url)


def select_option(elemid, option_text):
    """
    Select option from a drop down menu
    :param elemid: html element ID of menu
    :param option_text: name of option to select
    """
    elem = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, elemid)))
    drop = Select(elem)
    drop.select_by_visible_text(option_text)
    time.sleep(10)
    return


def change_window(link):
    """
    Switch driver control to the new window when clicking on a link that opens another window
    :param link: link that opens the new window when clicked
    """
    handles = driver.window_handles  # get previous number of windows
    link.click()
    WebDriverWait(driver, 60).until(EC.new_window_is_opened(handles))  # wait for new window up to 60 seconds
    new_window = driver.window_handles[1]
    driver.switch_to.window(new_window)  # change control to new window
    return


for cnpj in tqdm(lista_cnpjs):
    current_cnpj = cnpj.replace(".", "").replace("/", "").replace("-", "")  # remove punctuation from cnpj
    driver.get(url)
    try:
        field = driver.find_element_by_id('txtCNPJNome')
        field.clear()
        field.send_keys(current_cnpj)

        cont = driver.find_element_by_id('btnContinuar')
        cont.click()


        company_link = WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.ID, "dlCiasCdCVM__ctl1_Linkbutton1")))
        # create folder with company name
        company_name = unidecode.unidecode(str(company_link.text))
        path = "./companies_data/{}".format(company_name)
        # path = 'C:\\Users\\Owner\\Desktop\\DFs\\{}'.format(company_name)

        if os.path.exists(path):
            continue
        else:
            os.mkdir(path)

        company_link.click()

        periodo = WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.ID, "rdPeriodo")))
        periodo.click()

        dt_inicial = WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.ID, "txtDataIni")))
        dt_inicial.send_keys('04/01/2010')

        hr_inicial = WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.ID, "txtHoraIni")))
        hr_inicial.send_keys('00:00')

        dt_final = WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.ID, "txtDataFim")))
        dt_final.send_keys('22/05/2020')

        hr_final = WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.ID, "txtHoraFim")))
        hr_final.send_keys('00:00')
        hr_final.click()

        # Categoria
        categ_options = driver.find_element_by_id('cboCategoria')
        options = [i.get_attribute('text') for i in categ_options.find_elements_by_tag_name('option')]

        categ = driver.find_element_by_class_name('chosen-single')
        categ.click()
        categ = driver.find_element_by_id('cboCategoria_chosen_input')
        option = [i for i in options if i == 'DFP'][0]
        categ.send_keys(option)
        categ.send_keys(Keys.RETURN)

        # Consulta
        consulta = driver.find_element_by_id('btnConsulta')
        consulta.click()
        WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.ID, "grdDocumentos")))

        # selecting only documents with status "ATIVO"
        page = BeautifulSoup(driver.page_source, 'lxml')
        tableid = page.find('table', id='grdDocumentos')
        fulltable = pd.read_html(str(tableid), header=0)[0]
        ativos = fulltable[fulltable["Status"] == "Ativo"].index
        visus = driver.find_elements_by_id('VisualizarDocumento')
        visus = [visus[j] for j in ativos]

        print('----> current company: ' + company_name)

        for j in tqdm(visus):
            change_window(j)

            # choose the DF individual option
            select_option('cmbGrupo', 'DFs Individuais')

            # choose 'Balanço Patrimonial Ativo' and switch to iframe
            select_option('cmbQuadro', 'Balanço Patrimonial Ativo')
            WebDriverWait(driver, 500).until(EC.frame_to_be_available_and_switch_to_it(("iFrameFormulariosFilho")))

            # get 'Balanço Patrimonial Ativo' DataFrame
            page = BeautifulSoup(driver.page_source, 'lxml')
            table = page.find('table', id='ctl00_cphPopUp_tbDados')
            balanco_atv = pd.read_html(str(table), header=0)[0]

            # switch back to default content (out of iframe)
            driver.switch_to.default_content()

            # choose 'Balanço Patrimonial Passivo' and switch to iframe
            select_option('cmbQuadro', 'Balanço Patrimonial Passivo')
            WebDriverWait(driver, 500).until(EC.frame_to_be_available_and_switch_to_it("iFrameFormulariosFilho"))

            # get 'Balanço Patrimonial Passivo' DataFrame
            page = BeautifulSoup(driver.page_source, 'lxml')
            table = page.find('table', id='ctl00_cphPopUp_tbDados')
            balanco_pass = pd.read_html(str(table), header=0)[0]

            # switch back to default content (out of iframe)
            driver.switch_to.default_content()

            # choose 'Demonstração do Resultado' and switch to iframe
            select_option('cmbQuadro', 'Demonstração do Resultado')
            WebDriverWait(driver, 500).until(EC.frame_to_be_available_and_switch_to_it("iFrameFormulariosFilho"))

            # get 'Demonstração do Resultado' DataFrame
            page = BeautifulSoup(driver.page_source, 'lxml')
            table = page.find('table', id='ctl00_cphPopUp_tbDados')
            demonstra_res = pd.read_html(str(table), header=0)[0]

            # join DataFrames
            dfs_list = [balanco_atv, balanco_pass, demonstra_res]
            new_df = pd.concat(dfs_list)

            # get year (from first column name of DataFrame "balanco_atv") and create folder
            year = balanco_atv.columns[2][-4:]
            # path2 = path + '\\' + '{}'.format(year)
            path2 = path + '/' + '{}'.format(year)

            if os.path.exists(path2):
                continue
            else:
                os.mkdir(path2)

            # export DataFrame as csv file to folder path2:
            # csv_name = '\\' + '{}'.format(year) + '.csv'
            csv_name = '/' + '{}'.format(year) + '.csv'
            new_df.to_csv(path2 + csv_name)

            # close current window and change driver control to first window
            driver.close()
            first_window = driver.window_handles[0]
            driver.switch_to.window(first_window)

        driver.get(url)
    except:
        lista_cnpjs_non_cap.append(current_cnpj)
        driver.get(url)
        continue