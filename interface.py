import time
import os
import pandas as pd
import funcoes as my_fn
from appJar import gui

def _thread_tomados():
    ano = app.getOptionBox("Ano: ")
    mes = app.getOptionBox("Mês: ")
    dfs = app.openBox(title=None, dirName=None, fileTypes=None, asFile=False, parent=None, multiple=True,
                               mode='r')

    app.thread(my_fn._tomados(ano, mes, dfs))

def _thread_prestados():
    app.thread(my_fn._prestados())


# Criando a interface Gráfica
app = gui("The book is on the table")
app.setFont(10)

#app.addLabel("label_0_0", "1) Habilitar no Chrome: \"Perguntar onde salvar cada arquivo\nantes de fazer download\"",
             #row=0, column=0)

# app.addLabel("label_1_0", "Ano", row=1, column=0)
# app.addEntry("ano", row=1, column=1)
app.addLabelOptionBox("Ano: ", ["- Ano -", "2020", "2021"],  row=1, column=0)
# app.setEntryDefault("ano", "2020")

# app.addLabel("label_2_0", "Mês (no seguinte formato: 02.2020)", row=2, column=0)
app.addLabelOptionBox("Mês: ", ["- Mês -", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 'Julho', "Agosto",
                        "Setembro", "Outubro", "Novembro", "Dezembro" ],  row=2, column=0)
# app.setEntryDefault("mes", "02.2020")

# app.addLabel("label_3_0", "Tempo de espera para carregar páginas (não deixar em branco!)", row=3, column=0)
# app.addEntry("segundos", row=3, column=1)
# app.setEntryDefault("segundos", "Segundos")

app.addButton("NAO CLICAR - Prestados", _thread_prestados, row=4, column=0)
app.addButton("Tomados", _thread_tomados, row=4, column=1)

# start the GUI
app.go()

# Fazer o caminho onde salvar o tomados, criar uma lista no dicionario e talvez usar o cnpj como chave
# Preciso saber quais empresas terão as notas exportadas
