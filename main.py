from os import path, listdir
from docx2pdf import convert
from tkinter import filedialog
import PySimpleGUI as sg


layout = [
    [sg.Text('Coverter Word p/ PDF')],
    [sg.Button('Coverter')]
]

janela = sg.Window('Conversor', layout)

while True:
    event, values = janela.read()
    if event == sg.WIN_CLOSED:
        break

    try:
        # Abrir o gerenciador de tarefas
        pasta = filedialog.askdirectory()
        # Guarda o caminha da pasta
        caminho_pasta = path.join(pasta)
        # Guarda os nomes dos arquivos em uma lista
        lista_arquivos = listdir(pasta)

        if lista_arquivos and caminho_pasta:
            for item in range(len(lista_arquivos)):
                arquivo = lista_arquivos[item]
                convert(rf'{caminho_pasta}/{arquivo}')
            sg.popup('Conversão concluida...')
        else:
            sg.popup('Não existem arquivos na pasta selecionada...')
    except FileNotFoundError:
        pass




