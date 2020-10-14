import os
import re
import shutil
import time

import pandas as pd
import unidecode
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
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

url = 'https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/CiaAb/FormBuscaCiaAb.aspx?TipoConsult=c'
driver.get(url)

# set local download folder
# downloadpath = '/Users/renatoaranha/downloads'
downloadpath = '/Users/rlopesc/downloads'


def click_change_window(link):
    """
    Switch driver control to the new window when clicking on a link that opens another window
    :param link: link that opens the new window when clicked
    """
    handles = driver.window_handles  # get previous number of windows
    link.click()
    WebDriverWait(driver, 60).until(EC.new_window_is_opened(handles))  # wait for new window up to 60 seconds
    new_window = driver.window_handles[1]
    driver.switch_to.window(new_window)  # change control to new window
    time.sleep(5)
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

        company_link = driver.find_element_by_id('dlCiasCdCVM__ctl1_Linkbutton1')

        # create folder with company name
        company_name = unidecode.unidecode(str(company_link.text))
        path = "./companies_data/{}".format(company_name)
        # path = 'C:\\Users\\Owner\\Desktop\\DFs\\{}'.format(company_name)

        if os.path.exists(path):
            continue
        else:
            os.mkdir(path)

        print(' current company: ' + company_name)

        company_link.click()
        time.sleep(10)

        periodo = driver.find_element_by_id('rdPeriodo')
        periodo.click()
        time.sleep(5)

        dt_inicial = driver.find_element_by_id('txtDataIni')
        dt_inicial.clear()
        dt_inicial.send_keys('04/01/2010')

        hr_inicial = driver.find_element_by_id('txtHoraIni')
        hr_inicial.clear()
        hr_inicial.send_keys('00:00')

        dt_final = driver.find_element_by_id('txtDataFim')
        dt_final.clear()
        dt_final.send_keys('22/05/2020')

        dt_ref = driver.find_element_by_id('txtDataReferencia')
        dt_ref.clear()
        dt_ref.send_keys('01/01/2010')

        hr_final = driver.find_element_by_id('txtHoraFim')
        hr_final.send_keys('00:00')
        hr_final.click()

        # Categoria
        categ_options = driver.find_element_by_id('cboCategoria')
        options = [i.get_attribute('text') for i in categ_options.find_elements_by_tag_name('option')]

        categ = driver.find_element_by_class_name('chosen-single')
        categ.click()
        categ = driver.find_element_by_id('cboCategoria_chosen_input')
        option = [i for i in options if i == 'Formulário de Referência'][0]
        categ.send_keys(option)
        categ.send_keys(Keys.RETURN)

        # Consulta
        consulta = driver.find_element_by_id('btnConsulta')
        consulta.click()
        time.sleep(5)

        # selecting only documents with status "ATIVO" and Data Referência "2020"
        page = BeautifulSoup(driver.page_source, 'lxml')
        tableid = page.find('table', id='grdDocumentos')
        fulltable = pd.read_html(str(tableid), header=0)[0]
        fulltable["year"] = fulltable["Data Referência"].apply(lambda x: x[-4:])
        filtered_index = fulltable[fulltable["Status"] == "Ativo"].index
        visu = driver.find_element_by_id('VisualizarDocumento')

        click_change_window(visu)

        save_pdf = driver.find_element_by_id('btnGeraRelatorioPDF')
        save_pdf.click()
        time.sleep(5)

        filename1 = driver.find_element_by_id("hdnNumeroSequencialDocumento").get_attribute("value")
        filename2 = driver.find_element_by_id("hdnCodigoCvm").get_attribute("value")
        filename_prefix = filename1 + "_" + filename2

        wait = WebDriverWait(driver, 60).until(EC.frame_to_be_available_and_switch_to_it("iFrameModal"))

        time.sleep(5)
        page = BeautifulSoup(driver.page_source, 'lxml')
        boxes = page.find_all("a")

        # unselect all
        all_options = boxes[0].get("id")
        driver.find_element_by_id(all_options).click()

        # choose selection
        options_to_select = ['4. Fatores de risco', '5. Gerenciamento de riscos e controles internos']
        id_options_to_select = [i.get("id") for i in boxes if i.text in options_to_select]
        for i in id_options_to_select:
            driver.find_element_by_id(i).click()

        time.sleep(5)
        gerar_pdf = driver.find_element_by_id("btnConsulta")
        gerar_pdf.click()
        time.sleep(20)

        # renaming file and moving it to pertinet folder
        fname = re.sub('[^A-Za-z0-9]+', '', company_name) + ".pdf"
        for i in os.listdir(downloadpath):
            if filename_prefix in i:
                shutil.move(downloadpath + "/" + i,
                            os.path.join(path, fname))

        time.sleep(5)

        # close current window and change driver control to first window
        driver.close()
        first_window = driver.window_handles[0]
        driver.switch_to.window(first_window)

        driver.get(url)
    except:
        lista_cnpjs_non_cap.append(current_cnpj)
        driver.get(url)
        continue
