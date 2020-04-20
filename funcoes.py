import pyautogui
import time
import threading
import pandas as pd
import json
import os
import shutil

from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException, ElementNotInteractableException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait

def _pegar_ultimo_arquivo_modificado (search_dir):
    savedPath = os.getcwd()
    os.chdir(search_dir)
    files = filter(os.path.isfile, os.listdir(search_dir))
    files = [os.path.join(search_dir, f) for f in files] # add path to each file
    files.sort(key=lambda x: os.path.getmtime(x))
    os.chdir(savedPath)

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
    browser.switch_to.window(browser.window_handles[-1]) # Troca de janela
    browser.find_element_by_xpath("//*[contains(text(), 'Serviços Tomados')]").click()
    browser.find_element_by_xpath("//*[contains(text(), 'Declarar Notas Tomadas')]").click()
    browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/form/div[2]/ul/li[2]").click()
    browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/form/div[2]/div/div[2]/div[2]/input").click()
    time.sleep(1) # Esperar para o alerta ser gerado
    alert = browser.switch_to.alert # Aceitar o alerta
    qtd_notas = [int(s) for s in str.split(alert.text) if s.isdigit()] # qtd de notas existentes na página
    qtd_notas = qtd_notas[0]
    alert.accept()

    if qtd_notas > 0:
        mes = _converte_mes(mes)
        i = 1
        while i <= qtd_notas:
            aux = "/html/body/div[1]/section/div/div/div/form/div[2]/div/div[2]/div[1]/div[" +str(i)+ "]/div[4]/div[6]/select"
            aux = browser.find_element_by_xpath(aux).text
            mes_aux = aux[-2:]
            ano_aux = aux[:4]
            if mes_aux != mes or str(ano_aux) != str(ano): # se a nota não for da competência....
                # browser.find_element_by_xpath(
                  #  "/html/body/div[1]/section/div/div/div/form/div[2]/div/div[2]/div[1]/div[" + str(
                   #     i) + "]/div[9]/input").click() # .... clicar em remover
                pass
            i += 1

    #_clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[3]/input")

def _clicar_pelo_XPATH(browser, xpath):
    time.sleep(1)
    button = browser.find_element_by_xpath(xpath)
    ActionChains(browser).move_to_element(button).click(button).perform()

def _relatorio_tomados(nome_contabil, ano, mes, browser):
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe tomados\\"
    mes_aux = _converte_mes(mes)+"."+ano
    path += nome_contabil + "\\" + ano + "\\" + mes_aux
    try:
        os.mkdir(path)
        print (path)
    except FileNotFoundError:
        try:
            path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe tomados\\"
            path += nome_contabil + "\\" + ano
            os.mkdir(path)
            path += "\\" + mes_aux
            os.mkdir(path)
        except FileNotFoundError:
            return "Não existe pasta"
    except FileExistsError:
        pass

    _clicar_pelo_XPATH(browser, "//*[contains(text(), 'Livro Digital')]") # Clicar em Livro Digital
    browser.find_elements_by_xpath("//*[contains(text(), 'Livro Digital')]")[1].click() # Clicar em Livro Digital, de novo (nao posso usar a funcao para clicar por casa da ambiguidade)
    time.sleep(1)
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form[2]/input[3]") # Clicar em Gerar Novo Livro Digital
    time.sleep(1)
    browser.find_element_by_id("tipodec").find_element_by_xpath("//option[@value='T']").click() # Selecionar livros Tomados
    time.sleep(5) # Esperar carregar as informações dos livros
    elem = browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/div/form/div[4]/div/span")
    aux = str(elem.get_attribute('style'))[str(str(elem.get_attribute('style'))).find(":")+2:-1] # Descobrir se o livro tá aberto ou fechado

    if aux == "none": # Livro está aberto
        _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/div/form/div[6]/div/input")  # Clicar em Gerar
        browser.switch_to.alert.accept()  # Aceitar o alerta
        browser.switch_to.alert.accept()  # Aceitar o alerta

        time.sleep(1)
        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])
        time.sleep(1)
        browser.execute_script('window.print();')
        time.sleep(1)
        browser.close()

        time.sleep(1)
        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])
        time.sleep(1)
        browser.execute_script('window.print();')
        time.sleep(1)
        browser.close()
        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])

        origem = _pegar_ultimo_arquivo_modificado("C:\\Users\\"+os.getlogin()+"\Downloads")
        destino = path
        shutil.move(origem, destino)

    elif aux == "inline": # Livro está fechado
        _clicar_pelo_XPATH(browser, "//*[contains(text(), 'Relatórios')]") # Clicar em Relatórios
        _clicar_pelo_XPATH(browser, "//*[contains(text(), 'Relatório de Notas Aceitas')]") # Clicar em Relatórios
        competencia = str(ano) + '-' + str(_converte_mes(mes))
        time.sleep(1)
        browser.find_elements_by_xpath("//option[@value='" + competencia + "']")[1].click()  # Selecionar oeríodo final
        browser.find_element_by_xpath("//option[@value='" + competencia + "']").click()  # Selecionar período inicial
        time.sleep(1)
        _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[3]/div/input") # Clicar em Gerar Relatório

        time.sleep(1)
        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])
        time.sleep(1)
        browser.execute_script('window.print();')
        time.sleep(1)
        browser.close()
        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])

        origem = _pegar_ultimo_arquivo_modificado("C:\\Users\\"+os.getlogin()+"\Downloads")
        destino = path
        shutil.move(origem, destino)

def _pdf_notas_tomados(nome_contabil, ano, mes, browser):
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe tomados\\"
    path += nome_contabil + "\\" + ano + "\\" + mes

    #browser.switch_to.window(browser.window_handles[-1])

    competencia = str(ano)+'-'+str(_converte_mes(mes))
    browser.find_element_by_xpath("//*[contains(text(), 'Serviços Tomados')]").click()
    browser.find_element_by_xpath("//*[contains(text(), 'Notas Tomadas no Municipio')]").click()
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[1]/input") # Campo de Procurar
    browser.find_element_by_id("ListGeneratorFind_periodoTrib").find_element_by_xpath("//option[@value='"+competencia+"']").click() # Selecionar competencia
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[2]/input[8]") # Botão Procurar
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/table/thead/tr/th[1]/input") # Marcar caixinha
    try:
        _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/table/tbody[1]/tr/td/a") # Selecionar todos
    except:
        browser.close()
        browser.switch_to.window(browser.window_handles[0])
        browser.switch_to.window(browser.window_handles[-1])
        return False
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form[1]/input[3]") # PDF

    time.sleep(1)
    browser.switch_to.window(browser.window_handles[0])
    browser.switch_to.window(browser.window_handles[-1])
    time.sleep(1)
    browser.execute_script('window.print();')
    time.sleep(1)
    browser.close()
    browser.switch_to.window(browser.window_handles[-1])

    origem = _pegar_ultimo_arquivo_modificado("C:\\Users\\"+os.getlogin()+"\Downloads")
    destino = path
    shutil.move(origem, destino)

    return True

def _exportar_notas_fiscais_tomados(nome_contabil, ano, mes, browser):
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe tomados\\"
    path += nome_contabil

    competencia = str(ano)+'-'+str(_converte_mes(mes))
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[1]/input") # Campo de Procurar
    browser.find_element_by_id("ListGeneratorFind_periodoTrib").find_element_by_xpath("//option[@value='"+competencia+"']").click() # Selecionar competencia
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[2]/input[8]") # Botão Procurar _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/table/thead/tr/th[1]/input") # Marcar caixinha
    try:
        _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/table/tbody[1]/tr/td/a")  # Selecionar todos
    except:
        return
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/table/tbody[1]/tr/td/a") # Selecionar todos
    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form[1]/input[7]") # Exportar Xml Abrasf 2.02
    browser.close()
    browser.switch_to.window(browser.window_handles[-1])

    origem = _pegar_ultimo_arquivo_modificado("C:\\Users\\"+os.getlogin()+"\Downloads")
    destino = path
    shutil.move(origem, destino)

# def _relatorio_prestados(nome_fiscal, ano, mes, segundos, cnpj):
#     path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Fiscal\Impostos\Federais\Empresas\\"
#     path += nome_fiscal + "\\" + ano + "\\" + mes
#     try:
#         os.mkdir(path)
#     except FileNotFoundError:
#         try:
#             path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Fiscal\Impostos\Federais\Empresas\\"
#             path += nome_fiscal + "\\" + ano
#             os.mkdir(path)
#             path += "\\" + mes
#             if path[-1] == " ":
#                 path = path[:-1]
#             os.mkdir(path)
#         except FileNotFoundError:
#             return "nao existe pasta"
#     except FileExistsError:
#         pass
#
#     # Fechar download antigo
#     pyautogui.click(X=1421, Y=836)
#     # time.sleep(segundos-1)
#
#     # Clicar no campo de CNPJ
#     _limpar_campo(X=275, Y=407)
#     pyautogui.click(X=275, Y=407)
#     # time.sleep(segundos)
#
#     # Digitar CNPJ
#     kb.type_this(cnpj)
#     # time.sleep(segundos)
#
#     # Enter
#     kb.press_this(Key.enter)
#     time.sleep(segundos)
#
#     if pyautogui.locateOnScreen("imgs/nao_achou.png", confidence=0.8) != None:
#         return "nao_achou"
#
#     # Clique em Prestador
#     pyautogui.click(X=320, Y=677)
#     time.sleep(segundos)
#
#     if pyautogui.locateOnScreen("imgs/eh_comercio.png", confidence=0.8) != None:
#         # Fechar a aba
#         kb.press_this_with_LCTRL("w")
#         time.sleep(segundos)
#         return "eh_comercio"
#
#     # Clique em Livro Digital
#     pyautogui.click(X=1058, Y=137)
#     time.sleep(segundos - 1)
#
#     # Clique no outro Livro Digital
#     pyautogui.click(X=1056, Y=190)
#     time.sleep(segundos)
#
#     # Clique Gerar Novo Livro Digital
#     # Apertar END
#     kb.press_this(Key.end)
#     time.sleep(segundos)
#     pyautogui.click(pyautogui.locateCenterOnScreen("imgs/gerar_novo_livro_digital.png", confidence=0.9))
#     # pyautogui.click(316, 656)
#     time.sleep(segundos + 2)
#
#     if pyautogui.locateCenterOnScreen("imgs/livro_fechado.png", confidence=0.5) == None:
#
#         # print (pyautogui.locateCenterOnScreen("imgs/livro_fechado.png", confidence=0.5))
#         #
#         # print(nome_fiscal, " : livro não gerado")
#
#         # Clique em Gerar
#         pyautogui.click(X=255, Y=619)
#         time.sleep(segundos + 2)
#
#         # Clique em OK
#         pyautogui.click(X=890, Y=162)
#         time.sleep(segundos)
#
#         # Clique em OK, de novo
#         pyautogui.click(X=889, Y=183)
#         time.sleep(segundos)
#
#         # Salvar o relatorio de prestados
#         # Botao direito no centro da tela
#         mc.right_click_at(X=670, Y=505)
#         time.sleep(segundos)
#
#         # Clicar em imprimir
#         pyautogui.click(X=718, Y=629)
#         time.sleep(segundos)
#
#         # Clicar em Salvar
#         pyautogui.click(X=1063, Y=690)
#         time.sleep(segundos)
#
#         # Clique para digitar o caminho
#         pyautogui.click(X=1082, Y=168)
#         time.sleep(segundos)
#
#         # Digitar o caminho
#         kb.type_this(path)
#         time.sleep(segundos)
#
#         # Enter
#         kb.press_this(Key.enter)
#         time.sleep(segundos)
#
#         # Clicar em Salvar
#         pyautogui.click(X=1189, Y=821)
#         time.sleep(segundos)
#
#         if pyautogui.locateOnScreen("imgs/sobreescrita.png", confidence=0.8):
#             pyautogui.press("left")
#             pyautogui.press("enter")
#
#         # Fechar a aba
#         kb.press_this_with_LCTRL("w")
#         time.sleep(segundos)
#
#         return "relatorio prestados salvo"
#
#     else:
#
#         # Clicar em Relatórios
#         pyautogui.click(X=1279, Y=136)
#         time.sleep(segundos)
#
#         # Clicar em Relatórios de NFS'e Emitidas
#         pyautogui.click(X=1319, Y=189)
#         time.sleep(segundos)
#
#         # Clicar em Relatórios de NFS'e Emitidas
#         pyautogui.click(X=1319, Y=189)
#         time.sleep(segundos)
#
#         # Clicar em Selecionar o Período Inicial
#         pyautogui.click(X=368, Y=364)
#         time.sleep(segundos)
#
#         # Digitar período
#         periodo = ano + "-" + mes[:2]
#         kb.type_this(periodo)
#         time.sleep(segundos)
#
#         # Enter
#         kb.press_this(Key.enter)
#         time.sleep(segundos)
#
#         # Clicar em Selecionar o Período Final
#         pyautogui.click(X=1131, Y=364)
#         time.sleep(segundos)
#
#         # Digitar período
#         periodo = ano + "-" + mes[:2]
#         kb.type_this(periodo)
#         time.sleep(segundos)
#
#         # Enter
#         kb.press_this(Key.enter)
#         time.sleep(segundos)
#
#         # Clicar em Gerar Relatorio
#         pyautogui.click(X=311, Y=409)
#         time.sleep(segundos)
#
#         # Salvar o relatorio de prestados
#         # Botao direito no centro da tela
#         mc.right_click_at(X=670, Y=505)
#         time.sleep(segundos)
#
#         # Clicar em imprimir
#         pyautogui.click(X=718, Y=629)
#         time.sleep(segundos)
#
#         # Clicar em Salvar
#         pyautogui.click(X=1063, Y=690)
#         time.sleep(segundos)
#
#         # Clique para digitar o caminho
#         pyautogui.click(X=1082, Y=168)
#         time.sleep(segundos)
#
#         # Digitar o caminho
#         kb.type_this(path)
#         time.sleep(segundos)
#
#         # Enter
#         kb.press_this(Key.enter)
#         time.sleep(segundos)
#
#         # Clicar em Salvar
#         pyautogui.click(X=1189, Y=821)
#         time.sleep(segundos)
#
#         if pyautogui.locateOnScreen("imgs/sobreescrita.png", confidence=0.8):
#             pyautogui.press("left")
#             pyautogui.press("enter")
#
#         # Fechar a aba
#         kb.press_this_with_LCTRL("w")
#         time.sleep(segundos)
#
#         return "relatorio prestados salvo"
#
#
# def _exportar_notas_fiscais_prestados(nome_contabil, segundos):
#     path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe\\"
#     path += nome_contabil
#
#     # Clicar em Notas Eletrônicas
#     pyautogui.click(X=806, Y=139)
#     time.sleep(segundos)
#
#     # Clicar em Exportar Notas
#     # pyautogui.click(X=797,Y=435)
#     pyautogui.click(pyautogui.locateCenterOnScreen("imgs/exportar_notas.png", confidence=0.8))
#     time.sleep(segundos)
#
#     # Clicar em == Mes ==
#     pyautogui.click(X=512, Y=360)
#     time.sleep(segundos)
#
#     # Digitar o mes
#     kb.type_this(_converte_mes())
#     time.sleep(segundos)
#
#     # Enter
#     kb.press_this(Key.enter)
#     time.sleep(segundos)
#
#     # Clicar em == Ano ==
#     pyautogui.click(X=815, Y=359)
#     time.sleep(segundos)
#
#     # Digitar o ano
#     kb.type_this(str(app.getEntry("ano")))
#     time.sleep(segundos)
#
#     # Enter
#     kb.press_this(Key.enter)
#     time.sleep(segundos)
#
#     # Clicar em Exportar XML Abrasf 2.02
#     pyautogui.click(X=841, Y=405)
#     time.sleep(segundos + 1)
#
#     if pyautogui.locateCenterOnScreen("imgs/salvar_como.png", confidence=0.6) != None:
#         print(nome_contabil, ": salvar xml")
#
#         # Clique para digitar o caminho
#         pyautogui.click(X=1082, Y=168)
#         time.sleep(segundos)
#
#         # Digitar o caminho
#         kb.type_this(path)
#         time.sleep(segundos)
#
#         # Enter
#         kb.press_this(Key.enter)
#         time.sleep(segundos)
#
#         # Clicar em Salvar
#         pyautogui.click(X=1216, Y=819)
#         time.sleep(segundos)
#
#     # Fechar a aba
#     kb.press_this_with_LCTRL("w")
#     time.sleep(segundos)
#
#     return "xml de notas prestados salvo"

def _selecionar_certificado():
    time.sleep(1)
    pyautogui.press('down')
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
               #'savefile.default_directory': 'C:\\Users\\thiago.madeira\Desktop\\temp'
               }
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option('prefs', profile)
    chrome_options.add_argument('--kiosk-printing')
    ### END - Configuração para salvar o pdf ###

    browser = webdriver.Chrome(options=chrome_options)
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
        df = pd.read_excel(df, encoding='latin-1', sep=';', converters={'cnpj': lambda x: str(x)})
        for index in df.index:
            cnpj = df.at[index, df.columns[1]]
            nome_pasta = df.at[index, df.columns[0]]
            browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/form/div[2]/input[1]").clear()
            browser.find_element_by_xpath("/html/body/div[1]/section/div/div/div/form/div[2]/input[1]").send_keys(str(cnpj)) # Digitar CNPJ
            _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/div[2]/input[5]") # Clicar em Procurar

            # Clicar no primeiro resultado
            deu = False
            while deu == False:
                try:
                    _clicar_pelo_XPATH(browser, "/html/body/div[1]/section/div/div/div/form/table/tbody[3]/tr[1]/td[2]/a")
                    deu = True
                except StaleElementReferenceException:
                    time.sleep(1)
                except ElementNotInteractableException:
                    time.sleep(1)

            _aceitar_notas(ano, mes, browser)
            _relatorio_tomados(nome_pasta, ano, mes, browser)  # tomados
            tem_nota = _pdf_notas_tomados(nome_pasta, ano, mes, browser)  # tomados
            if tem_nota:
                _exportar_notas_fiscais_tomados(nome_pasta, ano, mes, browser)  # tomados

def _prestados():
    df = pd.read_excel("Taciane.xlsx", sep=";", converters={'C.N.P.J./C.P.F./C.E.I./C.A.E.P.F.': lambda x: str(x)})
    dictionary = {}

    file = open("relatorio_Taciane.txt", "a+")

    for i in df.index:
        cnpj = str(df.at[i, "C.N.P.J./C.P.F./C.E.I./C.A.E.P.F."])
        cnpj = cnpj.replace("\"", "")
        dictionary[cnpj] = [0, 0]
        dictionary[cnpj][0] = df.at[i, "Apelido"]  # nome contabil
        dictionary[cnpj][1] = df.at[i, "Apelido"]  # nome fiscal

    ano = app.getEntry("Ano")

    mes = app.getEntry("Mês")
    segundos = int(app.getEntry("segundos"))

    for item in dictionary:

        nome_contabil = dictionary[item][0]
        nome_fiscal = dictionary[item][1]
        cnpj = item

        aux = _relatorio_prestados(nome_fiscal, ano, mes, segundos, cnpj)  # prestados
        print(aux)
        if aux == "nao existe pasta" or aux == "eh_comercio" or aux == "nao_achou":
            file = open("relatorio_Taciane.txt", "a+")
            file.write(nome_contabil + " " + aux + "\n")
            file.close()
            continue
        elif aux == "relatorio prestados salvo":
            _exportar_notas_fiscais_prestados(nome_contabil, segundos)  # prestados
            file = open("relatorio_Taciane.txt", "a+")
            file.write(nome_contabil + " " + aux + "\n")
            file.close()