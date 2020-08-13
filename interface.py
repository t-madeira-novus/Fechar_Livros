import funcoes as my_fn
from appJar import gui


def _thread_tomados():
    ano = app.getOptionBox("Ano: ")
    mes = app.getOptionBox("Mês: ")
    dfs = app.openBox(title=None, dirName=None, fileTypes=None, asFile=False, parent=None, multiple=True, mode='r')
    app.thread(my_fn._tomados(ano, mes, dfs))


def _thread_prestados():
    ano = app.getOptionBox("Ano: ")
    mes = app.getOptionBox("Mês: ")
    dfs = app.openBox(title=None, dirName=None, fileTypes=None, asFile=False, parent=None, multiple=True,
                               mode='r')
    app.thread(my_fn._prestados(ano, mes, dfs))


def _ajuda(submenu):

    aux = 'Este sistema automatiza o processo de fechar livros de serviços prestados e tomados.\n\n' \
          'Escolha o mês, o ano e o tipo de serviço.\n\n '\
          'Seleciona a planilha que contém a lista de empresas e aguarde.\n\n' \
          'Importante: se parar de funcionar, alterar a planilha deixando como primeira empresa aquele em que ele ' \
          'parou de funcionar. Depois, reiniciar o programa.\n\n' \
          'Importante: mover o cursor do mouse enquanto a tarefa está sendo executada pode causar comportamentos imprevisíveis'
    if submenu == 'Como usar':
        app.infoBox('Como usar', aux)
    elif submenu == 'Versão':
        app.infoBox('Versão', 'Versão 1.1')

# Criando a interface Gráfica
app = gui("Fechar Livros")
app.setFont(10)
app.addMenuList("Ajuda", ["Como usar", "Versão"], _ajuda)

app.addLabelOptionBox("Ano: ", ["- Ano -", "2020", "2021"],  row=1, column=2)
app.addLabelOptionBox("Mês: ", ["- Mês -", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 'Julho', "Agosto",
                        "Setembro", "Outubro", "Novembro", "Dezembro" ],  row=1, column=0)
app.addButton("Prestados", _thread_prestados, row=4, column=0)
app.addButton("Tomados", _thread_tomados, row=4, column=1)

# start the GUI
app.go()
