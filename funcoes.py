import pyautogui
import time
import threading
import pandas as pd
import json
import os
import shutil

from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait

def _completa_cnpj(cnpj):
    print(cnpj)
    if len(str(cnpj)) < 14:
        diferenca = 14 - len(str(cnpj))
        cnpj_aux = cnpj
        cnpj = str(0) * diferenca
        cnpj += cnpj_aux
        return cnpj
    else:
        return str(cnpj)


def _exclui_arquivos(search_dir, string):
    """Excluir arquivos com a string no nome e que estejam no path"""

    saved_path = os.getcwd()
    try:
        os.chdir(search_dir)
    except:
        os.mkdir(search_dir)
        return

    files = filter(os.path.isfile, os.listdir(search_dir))
    files = [os.path.join(search_dir, f) for f in files]  # add path to each file
    indices = []  #lista pra guardar indices de arquivos que contem string

    i = 0
    for file in files:
        if string in file:
            indices.append(i)
        i += 1

    for i in indices:
        try:
            os.remove(files[i])
        except:
            pass

    os.chdir(saved_path)

def _esperar(segundos):
    i = 0
    while i < segundos:
        time.sleep(1)
        if i < segundos:
            time.sleep(1)
        i += 1


def _pegar_ultimo_arquivo_modificado(search_dir):
    saved_path = os.getcwd()
    os.chdir(search_dir)
    files = filter(os.path.isfile, os.listdir(search_dir))
    files = [os.path.join(search_dir, f) for f in files]  # add path to each file
    files.sort(key=lambda x: os.path.getmtime(x))
    os.chdir(saved_path)

    return files[-1]


def _converte_mes(mes):
    if str(mes) == "Janeiro":
        return "01"
    elif str(mes) == "Fevereiro":
        return "02"
    elif str(mes) == "Março":
        return "03"
    elif str(mes) == "Abril":
        return "04"
    elif str(mes) == "Maio":
        return "05"
    elif str(mes) == "Junho":
        return "06"
    elif str(mes) == "Julho":
        return "07"
    elif str(mes) == "Agosto":
        return "08"
    elif str(mes) == "Setembro":
        return "09"
    elif str(mes) == "Outubro":
        return "10"
    elif str(mes) == "Novembro":
        return "11"
    elif str(mes) == "Dezembro":
        return "12"


def _aceitar_notas(ano, mes, browser):
    _esperar(2)
    browser.switch_to.window(browser.window_handles[-1])  # Troca de janela
    _esperar(2)
    browser.find_element_by_xpath("//*[contains(text(), 'Serviços Tomados')]").click()
    browser.find_element_by_xpath("//*[contains(text(), 'Declarar Notas Tomadas')]").click()
    browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/form/div[2]/ul/li[2]").click()
    browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/form/div[2]/div/div[2]/div[2]/input").click()

    _esperar(4)
    alert = browser.switch_to.alert  # Mudar para o alerta
    qtd_notas = [int(s) for s in str.split(alert.text) if s.isdigit()]  # qtd de notas existentes na página
    qtd_notas = qtd_notas[0]
    alert.accept()  # Aceitar o alerta

    _esperar(2)

    if qtd_notas > 0:
        mes = _converte_mes(mes)

        i = qtd_notas
        while i >= 1:
            _esperar(2)
            aux = "/html/body/div[1]/section/div/div/div/form/div[2]/div/div[2]/div[1]/div[" + str(
                i) + "]/div[5]/div[1]/select"
            aux = browser.find_element_by_xpath(aux).text
            mes_aux = aux[-2:]
            ano_aux = aux[:4]
            if mes_aux != mes or str(ano_aux) != str(ano):  # se a nota não for da competência....

                print(mes_aux, mes, ano_aux, ano, i)
                browser.find_element_by_xpath(
                    "/html/body/div[1]/section/div/div/div/form/div[2]/div/div[2]/div[1]/div[" + str(
                        i) + "]/div[11]/input").click()  # .... clicar em remover
            i -= 1

    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[3]/input")
    try:
        WebDriverWait(browser, 3).until(ec.alert_is_present())
        browser.switch_to.alert.accept()  # Aceitar o alerta
    except:
        pass


def _clicar_pelo_XPATH(browser, xpath):
    time.sleep(1)
    button = browser.find_element_by_xpath(xpath)
    ActionChains(browser).move_to_element(button).click(button).perform()


def _relatorio_tomados(nome_pasta, ano, mes, browser):
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe tomados\\"
    mes_aux = _converte_mes(mes)+"."+ano
    path += nome_pasta + "\\" + ano + "\\" + mes_aux
    try:  # Tenta criar a pasta da competência
        os.mkdir(path)
    except FileNotFoundError:  # Caso não exista a pasta com o ano...
        try:  # ... tenta criar a pasta com o ano e depois a pasta da competência, dentro da pasta do ano.
            path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe tomados\\"
            path += nome_pasta + "\\" + ano
            os.mkdir(path)
            path += "\\" + mes_aux
            os.mkdir(path)
        except FileNotFoundError: # Se não conseguir criar a pasta do ano, é pq não existe a pasta da empresa
            return "Não existe pasta"
    except FileExistsError: # Se a pasta já existir, continua
        pass

    _clicar_pelo_XPATH(browser, "//*[contains(text(), 'Livro Digital')]") # Clicar em Livro Digital
    browser.find_elements_by_xpath("//*[contains(text(), 'Livro Digital')]")[1].click() # Clicar em Livro Digital, de novo (nao posso usar a funcao para clicar por casa da ambiguidade)
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form[2]/input[3]") # Clicar em Gerar Novo Livro Digital
    browser.find_element_by_id("tipodec").find_element_by_xpath("//option[@value='T']").click() # Selecionar livros Tomados
    if "FECHAR O LIVRO E GERAR A O DAM" in str(browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/div/form/div[3]/div/select").text):
        dam = True
    else:
        dam = False

    elem = browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/div/form/div[4]/div/span")
    _esperar(5)
    aux = str(elem.get_attribute('style'))[str(str(elem.get_attribute('style'))).find(":")+2:-1] # Descobrir se o livro tá aberto ou fechado

    # try: # Livro está aberto
    if aux == "none": # Livro está aberto
        _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/div/form/div[6]/div/input")  # Clicar em Gerar
        _esperar(5)

        try:
            browser.switch_to.alert.accept()  # Aceitar o alerta
            browser.switch_to.alert.accept()  # Aceitar o alerta
        except:
            pass

        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])

        _esperar(5)

        browser.execute_script('window.print();')
        _esperar(5)
        browser.close()

        origem = _pegar_ultimo_arquivo_modificado("E:\\Users\\thiago.madeira\\Downloads")#("C:\\Users\\"+os.getlogin()+"\Downloads")("C:\\Users\\" + os.getlogin() + "\Downloads")
        destino = path
        try:
            shutil.move(origem, destino)
        except:
            pass

        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])

        if dam == True:

            _esperar(5)

            browser.execute_script('window.print();')

            _esperar(5)

            browser.close()
            browser.switch_to.window(browser.window_handles[0])
            browser.switch_to.window(browser.window_handles[-1])

            origem = _pegar_ultimo_arquivo_modificado("E:\\Users\\thiago.madeira\\Downloads")#("C:\\Users\\"+os.getlogin()+"\Downloads")("C:\\Users\\"+os.getlogin()+"\Downloads")
            destino = path
            try:
                shutil.move(origem, destino)
            except:
                os.remove(origem)

        return "Arquivos Salvos"

    # except: # Livro está fechado
    elif aux == "inline": # Livro está fechado
        competencia = str(ano) + '-' + str(_converte_mes(mes))
        try:
            _clicar_pelo_XPATH(browser, "//*[contains(text(), 'Relatórios')]") # Clicar em Relatórios
            _clicar_pelo_XPATH(browser, "//*[contains(text(), 'Relatório de Notas Aceitas')]")  # Clicar em Relatórios de Notas Aceitas
            browser.find_elements_by_xpath("//option[@value='" + competencia + "']")[1].click()  # Selecionar oeríodo final
            browser.find_element_by_xpath("//option[@value='" + competencia + "']").click()  # Selecionar período inicial
            _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[3]/div/input") # Clicar em Gerar Relatório

        except: # Eh comercio
            _clicar_pelo_XPATH(browser, "/html/body/div[1]/div[1]/div[2]/div/div[2]/div/nav/ul/li[3]/a")  # Clicar em Livro Digital
            _clicar_pelo_XPATH(browser, "/html/body/div[1]/div[1]/div[2]/div/div[2]/div/nav/ul/li[3]/ul/li[1]/a")  # Clicar em Livro Digital de novo
            _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form[2]/div[1]/input")  # Clicar em Procurar
            browser.find_elements_by_xpath("//option[@value='" + competencia + "']")[1].click()  # Selecionar oeríodo final
            browser.find_element_by_xpath("//option[@value='" + competencia + "']").click()  # Selecionar período inicial
            _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form[2]/div[2]/input[1]")  # Clicar em Procurar
            _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form[2]/table/tbody[3]/tr[1]/td[2]/a")  # Clicar no primeiro resultado

        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])
        browser.execute_script('window.print();')
        browser.close()
        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])

        origem = _pegar_ultimo_arquivo_modificado("E:\\Users\\thiago.madeira\\Downloads")#("C:\\Users\\"+os.getlogin()+"\Downloads")("C:\\Users\\"+os.getlogin()+"\Downloads")
        destino = path
        try:
            shutil.move(origem, destino)
        except:
            os.remove(origem)
        return "Arquivos Salvos"

def _pdf_notas_tomados(nome_contabil, ano, mes, browser):
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe tomados\\"
    mes_aux = _converte_mes(mes) + "." + ano
    path += nome_contabil + "\\" + ano + "\\" + mes_aux

    #browser.switch_to.window(browser.window_handles[-1])

    competencia = str(ano)+'-'+str(_converte_mes(mes))
    browser.find_element_by_xpath("//*[contains(text(), 'Serviços Tomados')]").click()
    browser.find_element_by_xpath("//*[contains(text(), 'Notas Tomadas no Municipio')]").click()
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[1]/input") # Campo de Procurar
    browser.find_element_by_id("ListGeneratorFind_periodoTrib").find_element_by_xpath("//option[@value='"+competencia+"']").click() # Selecionar competencia
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[2]/input[8]") # Botão Procurar
    _esperar(2)
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/table/thead/tr/th[1]/input") # Marcar caixinha
    try:
        _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/table/tbody[1]/tr/td/a") # Selecionar todos
    except:
        browser.close()
        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])
        return False
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form[1]/input[3]") # PDF

    _esperar(5)

    browser.switch_to.window(browser.window_handles[0])
    browser.switch_to.window(browser.window_handles[-1])

    _esperar(5)
    
    browser.execute_script('window.print();')
    _esperar(2)
    browser.close()
    browser.switch_to.window(browser.window_handles[-1])

    origem = _pegar_ultimo_arquivo_modificado("E:\\Users\\thiago.madeira\\Downloads")#("C:\\Users\\"+os.getlogin()+"\Downloads")("C:\\Users\\"+os.getlogin()+"\Downloads")
    os.rename(origem, "E:\\Users\\thiago.madeira\\Downloads\\Notas.pdf")
    origem = _pegar_ultimo_arquivo_modificado("E:\\Users\\thiago.madeira\\Downloads")#("C:\\Users\\"+os.getlogin()+"\Downloads")("C:\\Users\\" + os.getlogin() + "\Downloads")
    destino = path
    try:
        shutil.move(origem, destino)
    except:
        os.remove(origem)

    return True

def _exportar_notas_fiscais_tomados(nome_contabil, ano, mes, browser):
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe tomados\\"
    path += nome_contabil

    competencia = str(ano)+'-'+str(_converte_mes(mes))
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[1]/input") # Campo de Procurar
    browser.find_element_by_id("ListGeneratorFind_periodoTrib").find_element_by_xpath("//option[@value='"+competencia+"']").click() # Selecionar competencia
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[2]/input[8]") # Botão Procurar
    _esperar(3)
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/table/thead/tr/th[1]/input") # Marcar caixinha
    try:
        _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/table/tbody[1]/tr/td/a")  # Selecionar todos
    except:
        return
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/table/tbody[1]/tr/td/a") # Selecionar todos
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form[1]/input[7]") # Exportar Xml Abrasf 2.02

    _esperar(3)

    browser.switch_to.window(browser.window_handles[-1])
    browser.close()
    browser.switch_to.window(browser.window_handles[-1])

    origem = _pegar_ultimo_arquivo_modificado("E:\\Users\\thiago.madeira\\Downloads")#("C:\\Users\\"+os.getlogin()+"\Downloads")("C:\\Users\\"+os.getlogin()+"\Downloads")
    destino = path
    _exclui_arquivos(destino, ".xml")
    try:
        shutil.move(origem, destino)
    except:
        os.remove(origem)

def _relatorio_prestados(nome_pasta, ano, mes, browser):
    browser.switch_to.window(browser.window_handles[0])
    browser.switch_to.window(browser.window_handles[-1])
    elem = None
    try: # Verificar se eh comercio
        elem = browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/p")
        if elem != None:
            browser.close()
            browser.switch_to.window(browser.window_handles[0])
            browser.switch_to.window(browser.window_handles[-1])
            return "Eh comércio"
    except:
        pass

    desconsolidada = False
    if browser.find_element_by_xpath("/html/body/div[1]/div[1]/div[3]/div/div[1]/div[2]").text == "DES CONSOLIDADA":
        desconsolidada = True

    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Fiscal\Impostos\Federais\Empresas\\"
    mes_aux = _converte_mes(mes)+"."+ano
    path += nome_pasta + "\\" + ano + "\\" + mes_aux
    try:
        os.mkdir(path)
    except FileNotFoundError:
        try:
            path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Fiscal\Impostos\Federais\Empresas\\"
            path += nome_pasta + "\\" + ano
            os.mkdir(path)
            path += "\\" + mes
            if path[-1] == " ":
                path = path[:-1]
            os.mkdir(path)
        except FileNotFoundError:
            return "Não existe pasta"
    except FileExistsError:
        pass


    _clicar_pelo_XPATH(browser, "//*[contains(text(), 'Livro Digital')]")  # Clicar em Livro Digital
    browser.find_elements_by_xpath("//*[contains(text(), 'Livro Digital')]")[1].click()  # Clicar em Livro Digital, de novo (nao posso usar a funcao para clicar por casa da ambiguidade)
    time.sleep(1)
    _clicar_pelo_XPATH(browser,
                       "/html/body/div[1]/section/div/div/div/form[2]/input[3]")  # Clicar em Gerar Novo Livro Digital
    time.sleep(1)
    elem = browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/div/form/div[4]/div/span")

    # Gambiarra para forçar espera
    i = 0
    while i < 5:
        if i < 5:
            time.sleep(1)
        i += 1

    aux = str(elem.get_attribute('style'))[
          str(str(elem.get_attribute('style'))).find(":") + 2:-1]  # Descobrir se o livro tá aberto ou fechado

    if aux == "none": # Livro está aberto
    #try:
        _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/div/form/div[6]/div/input")  # Clicar em Gerar

        WebDriverWait(browser, 3).until(ec.alert_is_present())
        browser.switch_to.alert.accept()  # Aceitar o alerta
        browser.switch_to.alert.accept()  # Aceitar o alerta

        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])
        browser.execute_script('window.print();')
        browser.close()
        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])

        if desconsolidada == True:
            browser.execute_script('window.print();')
            browser.close()
            browser.switch_to.window(browser.window_handles[0])
            browser.switch_to.window(browser.window_handles[-1])

        origem = _pegar_ultimo_arquivo_modificado("E:\\Users\\thiago.madeira\\Downloads")#("C:\\Users\\"+os.getlogin()+"\Downloads")("C:\\Users\\"+os.getlogin()+"\Downloads")
        destino = path
        try:
            shutil.move(origem, destino)
        except:
            pass

        return "Arquivos Salvos"

    elif aux == "inline": # Livro está fechado
    #except:
        _clicar_pelo_XPATH(browser, "//*[contains(text(), 'Relatórios')]") # Clicar em Relatórios
        # _clicar_pelo_XPATH(browser, "//*[contains(text(), 'Relatório de NFS’e Emitidas')]") # Clicar em Relatórios
        _clicar_pelo_XPATH(browser, "/html/body/div[1]/div[1]/div[2]/div/div[2]/div/nav/ul/li[8]/ul/li[1]/a")  # Clicar em Relatórios

        competencia = str(ano) + '-' + str(_converte_mes(mes))
        time.sleep(1)
        browser.find_elements_by_xpath("//option[@value='" + competencia + "']")[1].click()  # Selecionar oeríodo final
        browser.find_element_by_xpath("//option[@value='" + competencia + "']").click()  # Selecionar período inicial
        time.sleep(1)
        _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[2]/div/input") # Clicar em Gerar Relatório

        # Salvar pdf do relatório de notas
        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])
        browser.execute_script('window.print();')
        browser.close()
        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])

        origem = _pegar_ultimo_arquivo_modificado("E:\\Users\\thiago.madeira\\Downloads")#("C:\\Users\\"+os.getlogin()+"\Downloads")("C:\\Users\\"+os.getlogin()+"\Downloads")
        destino = path
        try:
            shutil.move(origem, destino)
        except:
            pass

        return "Arquivos Salvos"

def _exportar_notas_fiscais_prestados(nome_contabil, browser, ano, mes):
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe\\"
    path += nome_contabil

    _clicar_pelo_XPATH(browser, "//*[contains(text(), 'Notas Eletrônicas')]")  # Clicar em Notas Eletrônicas
    _clicar_pelo_XPATH(browser, "//*[contains(text(), 'Exportar Notas')]")  # Clicar em Exportar Notas
    mes = int(_converte_mes(mes))
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[1]/div[1]/div/select")
    browser.find_element_by_xpath("//option[@value='" + str(mes) + "']").click()  # Selecionar mes
    browser.find_element_by_xpath("//option[@value='" + str(ano) + "']").click()  # Selecionar ano
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[2]/div/input[4]")  # Clicar em Exportar XML Abrasf 2.0

    _esperar(2)

    browser.close()
    browser.switch_to.window(browser.window_handles[0])
    browser.switch_to.window(browser.window_handles[-1])

    origem = _pegar_ultimo_arquivo_modificado("E:\\Users\\thiago.madeira\\Downloads")#("C:\\Users\\"+os.getlogin()+"\Downloads")
    destino = path
    _exclui_arquivos(destino, ".xml")

    if ".xml" in str(origem):
        shutil.move(origem, destino)
        return "Arquivos Salvos"
    else:
        return "Sem Movimento"

def _selecionar_certificado():
    time.sleep(1)
    #pyautogui.press('down')
    pyautogui.press('enter')


def _tomados(ano, mes, dfs):
    ### BEGIN - Configuração para salvar o pdf ###
    appState = {
        "recentDestinations": [
            {
                "id": "Save as PDF",
                "origin": "local",
                "account": ""
            }
        ],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
    profile = {'printing.print_preview_sticky_settings.appState': json.dumps(appState),
               'safebrowsing.enabled': 'false',
               # 'savefile.default_directory': 'C:\\Users\\thiago.madeira\Desktop\\temp'
               }
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option('prefs', profile)
    chrome_options.add_argument('--kiosk-printing')
    ### END - Configuração para salvar o pdf ###

    browser = webdriver.Chrome(options=chrome_options,
                               executable_path= "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Sistemas Internos\Fechar Livros\chromedriver.exe")
    browser.implicitly_wait(1)
    thread1 = threading.Thread(target=_selecionar_certificado)
    thread1.start()
    browser.get('https://nfse.pjf.mg.gov.br/site/')
    browser.maximize_window()
    browser.get('https://nfse.pjf.mg.gov.br/contador/login.php')

    _clicar_pelo_XPATH(browser, "/html/body/div[1]/div/div[2]/div/div[2]/div/nav/ul/li[1]/a/span[1]") # Utilitários
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/div/div[2]/div/div[2]/div/nav/ul/li[1]/ul/li[5]/a") # Visitar Cliente
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[1]/input") # Campo de Procurar

    for df in dfs:
        df = pd.read_csv(df, encoding='latin-1', sep=';', converters={'CNPJ': lambda x: str(x)})
        for index in df.index:
            cnpj = _completa_cnpj(df.at[index, df.columns[1]])
            nome_pasta = df.at[index, df.columns[0]]
            if nome_pasta[-1] == " ": # Remover espaço em branco ao final
                nome_pasta = nome_pasta[:-1]
            browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/form/div[2]/input[1]").clear()
            browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/form/div[2]/input[1]").send_keys(str(cnpj)) # Digitar CNPJ
            _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[2]/input[5]") # Clicar em Procurar
            _esperar(2)
            # Clicar no primeiro resultado
            nao_cadastrado = False
            deu = False
            while deu == False:
                try:
                    _clicar_pelo_XPATH(browser,
                                       "/html/body/div[1]/section/div/div/div/form/table/tbody[3]/tr[1]/td[2]/a")
                    deu = True
                except:
                    _esperar(1)
                    try:
                        if browser.find_element_by_xpath(
                                "/html/body/div[1]/section/div/div/div/form/table/tbody[3]/tr/td").text == "Não foram encontrados resultados que atendem à sua pesquisa.":
                            nao_cadastrado = True
                            break
                    except:
                        pass

            if nao_cadastrado == True:
                _esperar(2)
                file = open("relatorio_tomados_" + mes + "-" + ano + ".txt", "a+")
                file.write(nome_pasta + " - " + "Não cadastrado no site" + "\n")
                file.close()

            else:
                _aceitar_notas(ano, mes, browser)
                aux = _relatorio_tomados(nome_pasta, ano, mes, browser)  # tomados
                tem_nota = _pdf_notas_tomados(nome_pasta, ano, mes, browser)  # tomados
                if tem_nota:
                    _exportar_notas_fiscais_tomados(nome_pasta, ano, mes, browser)  # tomados
                else:
                    aux += " (sem movimento)"

                file = open("relatorio_tomados_" + mes + "-" + ano + ".txt", "a+")
                file.write(nome_pasta + " - " + aux + "\n")
                file.close()

def _prestados(ano, mes, dfs):
    ### BEGIN - Configuração para salvar o pdf ###
    appState = {
        "recentDestinations": [
            {
                "id": "Save as PDF",
                "origin": "local",
                "account": ""
            }
        ],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
    profile = {'printing.print_preview_sticky_settings.appState': json.dumps(appState),
               'safebrowsing.enabled': 'false',
               #'savefile.default_directory': 'C:\\Users\\thiago.madeira\Desktop\\temp'
               }
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option('prefs', profile)
    chrome_options.add_argument('--kiosk-printing')
    ### END - Configuração para salvar o pdf ###

    browser = webdriver.Chrome(options=chrome_options,
                               executable_path="P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Sistemas Internos\Fechar Livros\chromedriver.exe")
    browser.implicitly_wait(1)
    thread1 = threading.Thread(target=_selecionar_certificado)
    thread1.start()
    browser.get('https://nfse.pjf.mg.gov.br/site/')
    browser.maximize_window()
    browser.get('https://nfse.pjf.mg.gov.br/contador/login.php')

    _clicar_pelo_XPATH(browser, "/html/body/div[1]/div/div[2]/div/div[2]/div/nav/ul/li[1]/a/span[1]") # Utilitários
    try:
        _clicar_pelo_XPATH(browser, "/html/body/div[4]/div/input")  # Aviso da Prefeitura sobre coronavirus
    except:
        pass
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/div/div[2]/div/div[2]/div/nav/ul/li[1]/ul/li[5]/a") # Visitar Cliente
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[1]/input") # Campo de Procurar

    for df in dfs:
        df = pd.read_csv(df, encoding='latin-1', sep=';', converters={'CNPJ': lambda x: str(x)})

        for index in df.index:

            cnpj = _completa_cnpj(df.at[index, df.columns[1]])
            nome_pasta = df.at[index, df.columns[0]]
            if nome_pasta[-1] == " ": # Remover espaço em branco ao final
                nome_pasta = nome_pasta[:-1]
            browser.switch_to.window(browser.window_handles[0])
            browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/form/div[2]/input[1]").clear()
            browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/form/div[2]/input[1]").send_keys(str(cnpj)) # Digitar CNPJ
            _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[2]/input[5]") # Clicar em Procurar

            file = open("relatorio_prestados_" + mes + "-" + ano + ".txt", "a+")

            # Clicar no primeiro resultado
            nao_cadastrado = False
            deu = False
            while deu == False:
                try:
                    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/table/tbody[3]/tr[1]/td[2]/a")
                    deu = True
                except :
                    time.sleep(1)
                    try:
                        if browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/form/table/tbody[3]/tr/td").text == "Não foram encontrados resultados que atendem à sua pesquisa.":
                            nao_cadastrado = True
                            break
                    except:
                        pass

            if nao_cadastrado == True:
                file.write(nome_pasta + " - " + "Não cadastrado no site" + "\n")
                file.close()

            else:
                aux = _relatorio_prestados(nome_pasta, ano, mes, browser)
                print (aux)

                if aux == "Arquivos Salvos":
                    aux = _exportar_notas_fiscais_prestados(nome_pasta, browser, ano, mes)
                    file.write(nome_pasta + " - " + aux + "\n")
                    file.close()
                else:
                    file.write(nome_pasta + " - " + aux + "\n")
                    file.close()