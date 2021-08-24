# cvm-web-scraper

CVM web data extraction from a list of companies

site for consulting companies information: https://cvmweb.cvm.gov.br/SWB/Sistemas/SCW/CPublica/CiaAb/FormBuscaCiaAb.aspx?TipoConsult=c

file containing the companies to gather information: _Empresas_listadas_B3-CORRETO.xlsx_

#### USAGE

After installing the required packages (following _requirements.txt_ file) in an isolated environment, execute:
* _cvm_crawler_main.py_ to get companies files per year (folder _companies_data_)
* _cvm_crawler_get_formulario.py_ to get companies pdfs (folder _companies_data_)
* _cvm_craeler_email.py_ to get companies e-mails (file _tabela_empresa_nomes.csv_)
