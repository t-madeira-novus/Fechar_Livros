import time
import os
import pandas as pd
import pyautogui

from appJar import gui

def _converte_mes():
    mes = app.getEntry("mes")
    mes = mes[:2]

    if str(mes) == "01":
        return "Jan"
    elif str(mes) == "02":
        return "Fev"
    elif str(mes) == "03":
        return "Mar"
    elif str(mes) == "04":
        return "Abr"
    elif str(mes) == "05":
        return "Mai"
    elif str(mes) == "06":
        return "Jun"
    elif str(mes) == "07":
        return "Jul"
    elif str(mes) == "08":
        return "Ago"
    elif str(mes) == "09":
        return "Set"
    elif str(mes) == "10":
        return "Out"
    elif str(mes) == "11":
        return "Nov"
    elif str(mes) == "12":
        return "Dez"


def _limpar_campo(X, Y):
    # Clique duplo no campo
    mc.double_left_click_at(X, Y)
    # Apertar Home
    kb.press_this(Key.home)
    # Apertar End segurando shift
    kb.press_this_with_LSHIFT(Key.end)


def _relatorio_tomados(cnpj, nome_contabil, ano, mes, segundos):
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe tomados\\"
    path += nome_contabil + "\\" + ano + "\\" + mes
    os.mkdir(path)

    # Clicar no campo de CNPJ
    mc.left_click_at(X=275, Y=407)
    time.sleep(segundos)

    # Digitar CNPJ
    kb.type_this(cnpj)
    time.sleep(segundos)

    # Enter
    kb.press_this(Key.enter)
    time.sleep(segundos)

    # Clique em Prestador
    mc.left_click_at(X=320, Y=677)
    time.sleep(segundos)

    # Clique em Livro Digital
    mc.left_click_at(X=1060, Y=138)
    time.sleep(segundos)

    # Clique no outro Livro Digital
    mc.left_click_at(X=1056, Y=190)
    time.sleep(segundos)

    # Clique Gerar Novo Livro Digital
    mc.left_click_at(X=300, Y=855)
    time.sleep(segundos)

    # Clique em tipo de declaracao (diferenca pro prestados)
    mc.left_click_at(X=331, Y=375)
    time.sleep(segundos)

    # Clique em tomados (diferenca pro prestados)
    mc.left_click_at(X=308, Y=416, )
    time.sleep(segundos)

    # Clique em Gerar
    mc.left_click_at(X=255, Y=619)
    time.sleep(segundos)

    # Clique em OK
    mc.left_click_at(X=886, Y=163)
    time.sleep(segundos)

    # Salvar DAM
    # Botao direito no centro da tela
    mc.right_click_at(X=670, Y=505)
    time.sleep(segundos)

    # Clicar em imprimir
    mc.left_click_at(X=718, Y=629)
    time.sleep(segundos)

    # Clicar em Salvar
    mc.left_click_at(X=1063, Y=690)
    time.sleep(segundos)

    # Clique para digitar o caminho
    mc.left_click_at(X=1034, Y=170)
    time.sleep(segundos)

    # Digitar o caminho
    kb.type_this(path)
    time.sleep(segundos)

    # Enter
    kb.press_this(Key.enter)
    time.sleep(segundos)

    # Clicar em Salvar
    mc.left_click_at(X=1189, Y=821)
    time.sleep(segundos)

    # Fechar a aba
    kb.press_this_with_LCTRL("w")
    time.sleep(segundos)

    # Salvar relatorio
    # Botao direito no centro da tela
    mc.right_click_at(X=670, Y=505)
    time.sleep(segundos)

    # Clicar em imprimir
    mc.left_click_at(X=718, Y=629)
    time.sleep(segundos)

    # Clicar em Salvar
    mc.left_click_at(X=1063, Y=690)
    time.sleep(segundos)

    # Clique para digitar o caminho
    mc.left_click_at(X=1034, Y=170)
    time.sleep(segundos)

    # Digitar o caminho
    kb.type_this(path)
    time.sleep(segundos)

    # Enter
    kb.press_this(Key.enter)
    time.sleep(segundos)

    # Clicar em Salvar
    mc.left_click_at(X=1189, Y=821)
    time.sleep(segundos)

    # Fechar a aba
    kb.press_this_with_LCTRL("w")
    time.sleep(segundos)


def _pdf_notas_tomados(nome_contabil, ano, mes, segundos):
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe tomados\\"
    path += nome_contabil + "\\" + ano + "\\" + mes

    # Clicar em servicos tomados
    mc.left_click_at(X=937, Y=139)
    time.sleep(segundos)

    # Clicar em notas tomadas no municipio
    mc.left_click_at(X=919, Y=258)
    time.sleep(segundos)

    # Clicar em procurar
    mc.left_click_at(X=268, Y=351)
    time.sleep(segundos)

    # Clicar em competencia
    mc.left_click_at(X=266, Y=405)
    time.sleep(segundos)

    # Digitar a competencia
    competencia = ano + "-" + mes[:2]
    kb.type_this(competencia)
    time.sleep(segundos)

    # Enter
    kb.press_this(Key.enter)
    time.sleep(segundos)

    # Clicar em Procurar
    mc.left_click_at(X=264, Y=789)
    time.sleep(segundos)

    # Apertar page down
    kb.press_this(Key.page_down)
    time.sleep(segundos)

    # Clicar na caixinha
    mc.left_click_at(X=243, Y=355)
    time.sleep(segundos)

    # Clicar em selecionar todos
    mc.left_click_at(X=1069, Y=398)
    time.sleep(segundos)

    # Clicar em PDF
    mc.left_click_at(X=301, Y=684)
    time.sleep(segundos + 3)

    # Botao direito no centro da tela
    mc.right_click_at(X=670, Y=505)
    time.sleep(segundos)

    # Clicar em imprimir
    mc.left_click_at(X=718, Y=629)
    time.sleep(segundos)

    # Clicar em Salvar
    mc.left_click_at(X=1063, Y=690)
    time.sleep(segundos)

    # Clique para digitar o caminho
    mc.left_click_at(X=1004, Y=169)
    time.sleep(segundos)

    # Digitar o caminho
    kb.type_this(path)
    time.sleep(segundos)

    # Enter
    kb.press_this(Key.enter)
    time.sleep(segundos)

    # Limpar campo do nome do arquivo
    _limpar_campo(X=1118, Y=749)

    # Digitar nome do arquivo
    kb.type_this("Notas")
    time.sleep(segundos)

    # Clicar em Salvar
    mc.left_click_at(X=1194, Y=819)
    time.sleep(segundos)

    # Fechar a aba
    kb.press_this_with_LCTRL("w")
    time.sleep(segundos)


def _exportar_notas_fiscais_tomados(nome_contabil, segundos):
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe tomados\\"
    path += nome_contabil

    # Clicar em procurar
    mc.left_click_at(X=268, Y=351)
    time.sleep(segundos)

    # Clicar em competencia
    mc.left_click_at(X=266, Y=405)
    time.sleep(segundos)

    # Digitar a competencia
    competencia = ano + "-" + mes[:2]
    kb.type_this(competencia)
    time.sleep(segundos)

    # Enter
    kb.press_this(Key.enter)
    time.sleep(segundos)

    # Clicar em Procurar
    mc.left_click_at(X=264, Y=789)
    time.sleep(segundos)

    # Apertar page down
    kb.press_this(Key.page_down)
    time.sleep(segundos)

    # Clicar na caixinha
    mc.left_click_at(X=243, Y=355)
    time.sleep(segundos)

    # Clicar em selecionar todos
    mc.left_click_at(X=1069, Y=398)
    time.sleep(segundos)

    # Clicar em Exportar XML Abrasf 2.02
    mc.left_click_at(X=905, Y=685)
    time.sleep(segundos)

    # Clique para digitar o caminho
    mc.left_click_at(X=1004, Y=169)
    time.sleep(segundos)

    # Digitar o caminho
    kb.type_this(path)
    time.sleep(segundos)

    # Enter
    kb.press_this(Key.enter)
    time.sleep(segundos)

    # Clicar em Salvar
    mc.left_click_at(X=1182, Y=824)
    time.sleep(segundos)


def _relatorio_prestados(nome_fiscal, ano, mes, segundos, cnpj):
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Fiscal\Impostos\Federais\Empresas\\"
    path += nome_fiscal + "\\" + ano + "\\" + mes
    try:
        os.mkdir(path)
    except FileNotFoundError:
        try:
            path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Fiscal\Impostos\Federais\Empresas\\"
            path += nome_fiscal + "\\" + ano
            os.mkdir(path)
            path += "\\" + mes
            if path[-1] == " ":
                path = path[:-1]
            os.mkdir(path)
        except FileNotFoundError:
            return "nao existe pasta"
    except FileExistsError:
        pass

    # Fechar download antigo
    pyautogui.click(X=1421, Y=836)
    # time.sleep(segundos-1)

    # Clicar no campo de CNPJ
    _limpar_campo(X=275, Y=407)
    pyautogui.click(X=275, Y=407)
    # time.sleep(segundos)

    # Digitar CNPJ
    kb.type_this(cnpj)
    # time.sleep(segundos)

    # Enter
    kb.press_this(Key.enter)
    time.sleep(segundos)

    if pyautogui.locateOnScreen("imgs/nao_achou.png", confidence=0.8) != None:
        return "nao_achou"

    # Clique em Prestador
    pyautogui.click(X=320, Y=677)
    time.sleep(segundos)

    if pyautogui.locateOnScreen("imgs/eh_comercio.png", confidence=0.8) != None:
        # Fechar a aba
        kb.press_this_with_LCTRL("w")
        time.sleep(segundos)
        return "eh_comercio"

    # Clique em Livro Digital
    pyautogui.click(X=1058, Y=137)
    time.sleep(segundos - 1)

    # Clique no outro Livro Digital
    pyautogui.click(X=1056, Y=190)
    time.sleep(segundos)

    # Clique Gerar Novo Livro Digital
    # Apertar END
    kb.press_this(Key.end)
    time.sleep(segundos)
    pyautogui.click(pyautogui.locateCenterOnScreen("imgs/gerar_novo_livro_digital.png", confidence=0.9))
    # pyautogui.click(316, 656)
    time.sleep(segundos + 2)

    if pyautogui.locateCenterOnScreen("imgs/livro_fechado.png", confidence=0.5) == None:

        # print (pyautogui.locateCenterOnScreen("imgs/livro_fechado.png", confidence=0.5))
        #
        # print(nome_fiscal, " : livro não gerado")

        # Clique em Gerar
        pyautogui.click(X=255, Y=619)
        time.sleep(segundos + 2)

        # Clique em OK
        pyautogui.click(X=890, Y=162)
        time.sleep(segundos)

        # Clique em OK, de novo
        pyautogui.click(X=889, Y=183)
        time.sleep(segundos)

        # Salvar o relatorio de prestados
        # Botao direito no centro da tela
        mc.right_click_at(X=670, Y=505)
        time.sleep(segundos)

        # Clicar em imprimir
        pyautogui.click(X=718, Y=629)
        time.sleep(segundos)

        # Clicar em Salvar
        pyautogui.click(X=1063, Y=690)
        time.sleep(segundos)

        # Clique para digitar o caminho
        pyautogui.click(X=1082, Y=168)
        time.sleep(segundos)

        # Digitar o caminho
        kb.type_this(path)
        time.sleep(segundos)

        # Enter
        kb.press_this(Key.enter)
        time.sleep(segundos)

        # Clicar em Salvar
        pyautogui.click(X=1189, Y=821)
        time.sleep(segundos)

        if pyautogui.locateOnScreen("imgs/sobreescrita.png", confidence=0.8):
            pyautogui.press("left")
            pyautogui.press("enter")

        # Fechar a aba
        kb.press_this_with_LCTRL("w")
        time.sleep(segundos)

        return "relatorio prestados salvo"

    else:

        # Clicar em Relatórios
        pyautogui.click(X=1279, Y=136)
        time.sleep(segundos)

        # Clicar em Relatórios de NFS'e Emitidas
        pyautogui.click(X=1319, Y=189)
        time.sleep(segundos)

        # Clicar em Relatórios de NFS'e Emitidas
        pyautogui.click(X=1319, Y=189)
        time.sleep(segundos)

        # Clicar em Selecionar o Período Inicial
        pyautogui.click(X=368, Y=364)
        time.sleep(segundos)

        # Digitar período
        periodo = ano + "-" + mes[:2]
        kb.type_this(periodo)
        time.sleep(segundos)

        # Enter
        kb.press_this(Key.enter)
        time.sleep(segundos)

        # Clicar em Selecionar o Período Final
        pyautogui.click(X=1131, Y=364)
        time.sleep(segundos)

        # Digitar período
        periodo = ano + "-" + mes[:2]
        kb.type_this(periodo)
        time.sleep(segundos)

        # Enter
        kb.press_this(Key.enter)
        time.sleep(segundos)

        # Clicar em Gerar Relatorio
        pyautogui.click(X=311, Y=409)
        time.sleep(segundos)

        # Salvar o relatorio de prestados
        # Botao direito no centro da tela
        mc.right_click_at(X=670, Y=505)
        time.sleep(segundos)

        # Clicar em imprimir
        pyautogui.click(X=718, Y=629)
        time.sleep(segundos)

        # Clicar em Salvar
        pyautogui.click(X=1063, Y=690)
        time.sleep(segundos)

        # Clique para digitar o caminho
        pyautogui.click(X=1082, Y=168)
        time.sleep(segundos)

        # Digitar o caminho
        kb.type_this(path)
        time.sleep(segundos)

        # Enter
        kb.press_this(Key.enter)
        time.sleep(segundos)

        # Clicar em Salvar
        pyautogui.click(X=1189, Y=821)
        time.sleep(segundos)

        if pyautogui.locateOnScreen("imgs/sobreescrita.png", confidence=0.8):
            pyautogui.press("left")
            pyautogui.press("enter")

        # Fechar a aba
        kb.press_this_with_LCTRL("w")
        time.sleep(segundos)

        return "relatorio prestados salvo"


def _exportar_notas_fiscais_prestados(nome_contabil, segundos):
    path = "P:\documentos\OneDrive - Novus Contabilidade\Doc Compartilhado\Contábil\\NFSe\\"
    path += nome_contabil

    # Clicar em Notas Eletrônicas
    pyautogui.click(X=806, Y=139)
    time.sleep(segundos)

    # Clicar em Exportar Notas
    # pyautogui.click(X=797,Y=435)
    pyautogui.click(pyautogui.locateCenterOnScreen("imgs/exportar_notas.png", confidence=0.8))
    time.sleep(segundos)

    # Clicar em == Mes ==
    pyautogui.click(X=512, Y=360)
    time.sleep(segundos)

    # Digitar o mes
    kb.type_this(_converte_mes())
    time.sleep(segundos)

    # Enter
    kb.press_this(Key.enter)
    time.sleep(segundos)

    # Clicar em == Ano ==
    pyautogui.click(X=815, Y=359)
    time.sleep(segundos)

    # Digitar o ano
    kb.type_this(str(app.getEntry("ano")))
    time.sleep(segundos)

    # Enter
    kb.press_this(Key.enter)
    time.sleep(segundos)

    # Clicar em Exportar XML Abrasf 2.02
    pyautogui.click(X=841, Y=405)
    time.sleep(segundos + 1)

    if pyautogui.locateCenterOnScreen("imgs/salvar_como.png", confidence=0.6) != None:
        print(nome_contabil, ": salvar xml")

        # Clique para digitar o caminho
        pyautogui.click(X=1082, Y=168)
        time.sleep(segundos)

        # Digitar o caminho
        kb.type_this(path)
        time.sleep(segundos)

        # Enter
        kb.press_this(Key.enter)
        time.sleep(segundos)

        # Clicar em Salvar
        pyautogui.click(X=1216, Y=819)
        time.sleep(segundos)

    # Fechar a aba
    kb.press_this_with_LCTRL("w")
    time.sleep(segundos)

    return "xml de notas prestados salvo"


def _thread_tomados():
    app.thread(_tomados())


def _tomados():
    dictionary = {}
    # dictionary["25510747000110"] = [0,0]
    # dictionary["25510747000110"][0] = "541-NOVUS CORRETORA" # nome contabil
    # dictionary["25510747000110"][1] = "Novus Corretora"  # nome fiscal

    ano = app.getEntry("ano")
    mes = app.getEntry("mes")
    segundos = int(app.getEntry("segundos"))

    for item in dictionary:
        nome_contabil = dictionary[item][0]
        nome_fiscal = dictionary[item][1]
        cnpj = stritem

        _relatorio_tomados(cnpj, nome_contabil, ano, mes, segundos)  # tomados
        _pdf_notas_tomados(nome_contabil, ano, mes, segundos)  # tomados
        _exportar_notas_fiscais_tomados(nome_contabil, segundos)  # tomados


def _thread_prestados():
    app.thread(_prestados())


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

    ano = app.getEntry("ano")

    mes = app.getEntry("mes")
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


# Criando a interface Gráfica
app = gui("The book is on the table")
app.setFont(10)

app.addLabel("label_0_0", "1) Habilitar no Chrome: \"Perguntar onde salvar cada arquivo antes de fazer download\"",
             row=0, column=0)

app.addLabel("label_1_0", "Ano", row=1, column=0)
app.addEntry("ano", row=1, column=1)
app.setEntryDefault("ano", "2020")

app.addLabel("label_2_0", "Mês (no seguinte formato: 02.2020)", row=2, column=0)
app.addEntry("mes", row=2, column=1)
app.setEntryDefault("mes", "02.2020")

app.addLabel("label_3_0", "Tempo de espera para carregar páginas (não deixar em branco!)", row=3, column=0)
app.addEntry("segundos", row=3, column=1)
app.setEntryDefault("segundos", "Segundos")

app.addButton("Prestados", _thread_prestados, row=4, column=0)
app.addButton("NAO CLICAR - Tomados", _thread_tomados, row=4, column=1)

# start the GUI
app.go()

# Fazer o caminho onde salvar o tomados, criar uma lista no dicionario e talvez usar o cnpj como chave
# Preciso saber quais empresas terão as notas exportadas
