from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import openpyxl
import pathlib
from openpyxl import Workbook, workbook
from time import time
from datetime import datetime

chromedriver_path = r'C:\Users\compras.mg\OneDrive\Área de Trabalho\Projeto Fipe\ConsultaTbFipe\chromedriver.exe'

# Abrir a planilha Excel
workbook = openpyxl.load_workbook(r"C:\Users\compras.mg\OneDrive\Área de Trabalho\Projeto Fipe\ConsultaTbFipe\dadosFipe.xlsx")
sheet = workbook.active

# Ler os códigos FIPE e anos da planilha
codigos_fipe = [cell.value for cell in sheet['A']]
anos = [cell.value for cell in sheet['B']]

# Configurar o serviço do Chrome
service = Service(chromedriver_path)

# Criar uma instância do driver do Chrome
driver = webdriver.Chrome(service=service)

sleep(2)

driver.get('https://veiculos.fipe.org.br/')
sleep(2)

# Localizar o elemento usando XPath
carros_pequenos = driver.find_element(By.XPATH,
                                      '//*[@id="front"]/div[1]/div[2]/ul/li[1]')
carros_pequenos.click()
sleep(3)
pesquisa_codigo_fipe = driver.find_element(By.XPATH,
                                           '//*[@id="front"]/div[1]/div[2]/ul/li[1]/div/nav/ul/li[2]/a')
pesquisa_codigo_fipe.click()

sleep(2)

pesquisa_codigo_fipe = driver.find_element(By.XPATH,
                                           '//*[@id="front"]/div[1]/div[2]/ul/li[1]/div/nav/ul/li[2]/a')
pesquisa_codigo_fipe.click()
'''
# Seleciona o campo onde tem os messes e anos de referencia
pesquisa_mes_ref = driver.find_element(By.XPATH,
                                           '//*[@id="selectTabelaReferenciacarroCodigoFipe_chosen"]/a/div/b')
pesquisa_mes_ref.click()

#indice do mês de referencia
indice_mes_ref = [9] # Mês atual sempre [0] e mês anterior retroagir esse número
# Seleciona o mes e ano de referencia
mes_ref = driver.find_element(By.XPATH, f'//*[@id="selectTabelaReferenciacarroCodigoFipe_chosen"]/div/ul/li[0]')
mes_ref.click()                       #//li[text()="dezembro/2022"]
'''
contPross = 0
cont_time = 0
# Obtém o tempo de início
start_time = time()


# Crie um arquivo Excel
result_workbook = Workbook()
result_sheet = result_workbook.active
result_sheet['A1'] = "Mês de Referência"
result_sheet['B1'] = "Código FIPE Result"
result_sheet['C1'] = "Marca"
result_sheet['D1'] = "Modelo"
result_sheet['E1'] = "Ano Modelo"
result_sheet['F1'] = "Preço Médio"  
    


# Loop sobre os códigos FIPE e anos
for codigo, ano in zip(codigos_fipe, anos):
    # Localize o elemento de entrada de código FIPE
    codigo_fipe = driver.find_element(By.XPATH,
                                      '//*[@id="selectCodigocarroCodigoFipe"]')
    codigo_fipe.clear()
    codigo_fipe.send_keys(codigo)
    sleep(2)

    # Clique no elemento de seleção personalizado para abrir as opções
    menu_ano = driver.find_element(By.XPATH, '//*[@id="selectCodigoAnocarroCodigoFipe_chosen"]/a')
    menu_ano.click()

    # Aguarde um momento para as opções serem carregadas
    sleep(2)

    # Selecione a opção correspondente ao ano
    opcao_ano = driver.find_element(By.XPATH, f'//li[text()="{ano}"]')
    opcao_ano.click()
    sleep(2)

    # Clicando no botão para pesquisar
    botao_pesquiser = driver.find_element(By.XPATH, '//*[@id="buttonPesquisarcarroPorCodigoFipe"]')
    botao_pesquiser.click()
    sleep(3)

    # Armazena os resultados em lista

    mes_referencia = []
    # localizar o campo mês referência
    mes = driver.find_element(By.XPATH,
                              '//*[@id="resultadocarroCodigoFipe"]/table/tbody/tr[1]/td[2]/p')
    sleep(2)
    # Extrair o mês de referência do elemento 'mes' e adicioná-lo à lista 'mes_referencia'
    mes_referencia.append(mes.get_attribute("textContent").strip())

    codigo_result = []
    # localiza o campo com o código FIPE
    fipe = driver.find_element(By.XPATH, '//*[@id="resultadocarroCodigoFipe"]/table/tbody/tr[2]/td[2]/p')

    # Extraindo o código do elemento e adicionando na lista 'codigo_result'
    codigo_result.append(fipe.get_attribute("textContent").strip())

    marca = []
    montadora = driver.find_element(By.XPATH, '//*[@id="resultadocarroCodigoFipe"]/table/tbody/tr[3]/td[2]/p')

    # Extraindo a montadora do elemento e adicionando na lista 'marca'
    marca.append(montadora.get_attribute("textContent").strip())

    modelo = []
    mod = driver.find_element(By.XPATH, '//*[@id="resultadocarroCodigoFipe"]/table/tbody/tr[4]/td[2]/p')

    # Extraindo o modelo do veículo e adicionando na lista 'modelo'
    modelo.append(mod.get_attribute("textContent").strip())

    ano_modelo = []
    ano_mod = driver.find_element(By.XPATH, '//*[@id="resultadocarroCodigoFipe"]/table/tbody/tr[5]/td[2]/p')

    # Extraindo o ano modelo e adicionando na lista 'ano_modelo'
    ano_modelo.append(ano_mod.get_attribute("textContent").strip())

    preco_medio = []
    preco = driver.find_element(By.XPATH, '//*[@id="resultadocarroCodigoFipe"]/table/tbody/tr[8]/td[2]/p')

    # Extraindo o preço médio e adicionando na lista 'preco_medio'
    preco_medio.append(preco.get_attribute("textContent").strip())

    # vLimpav o campo 'código FIPE'
    limpar_pesquisa = driver.find_element(By.XPATH, '//*[@id="buttonLimparPesquisarcarroPorCodigoFipe"]/a')
    limpar_pesquisa.click()
    sleep(2)


    # Obtém o tempo atual
    current_time = time()

    # Calcula o tempo gasto desde o início do loop
    elapsed_time = current_time - start_time

    # Formata o tempo decorrido em horas, minutos e segundos
    formatted_time = datetime.utcfromtimestamp(elapsed_time).strftime('%H:%M:%S')

    # Atualiza a variável cont_time
    cont_time = formatted_time
    contPross = contPross+1
    
    print(f'feito {contPross} realizado por ultimo {codigo_result} {ano_modelo} tempo gasto para consulta {cont_time}')
    # Escreva os dados na planilha de resultados
    result_sheet.append([mes_referencia[0], codigo_result[0], marca[0], modelo[0], ano_modelo[0], preco_medio[0]])
    
    
    result_workbook.save(r"C:\Users\compras.mg\OneDrive\Área de Trabalho\Projeto Fipe\ConsultaTbFipe\resultadoFinal.xlsx")

# Salve o arquivo Excel de resultados
result_workbook.save(r"C:\Users\compras.mg\OneDrive\Área de Trabalho\Projeto Fipe\ConsultaTbFipe\resultadoFinal.xlsx")

print("Processo realizado com sucesso!")
# Encerrar o driver
driver.quit()
