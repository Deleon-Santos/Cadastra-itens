#estou desenvolvondo um novo projeto utilizando tabela em pysiplegui

import PySimpleGUI as sg

lista=[["1","101","78900000101","acucar uniao de 1kg","1","3.90"]]
titulos=["N°","Item","EAN","Descrição","QTD","Preço"]

layout= [
    [sg.Table(values = lista,headings=titulos,max_col_width=10,auto_size_columns=True, display_row_numbers=True,justification = "right", num_rows=30,
              key="-TABELA-",row_height=20)],
    [sg.Button("SAIR")],]

window = sg.Window("Tabela com valores", layout,size=(600,400))
evento, valor = window.read()
while True:
    if evento in (sg.WIN_CLOSED,"SAIR"):
        break