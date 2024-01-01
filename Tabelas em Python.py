import PySimpleGUI as sg
import pandas as pd
lista = [
    
]
titulos = ["N°", "Item", "EAN", "Descrição do Produto", "QTD", "Preço"]

tabela = [
    [sg.Table(values=lista, headings=titulos, max_col_width=10, auto_size_columns=True,
              display_row_numbers=True, justification="right", num_rows=30, key="-TABELA-", row_height=12)],
]

entradas = [
    [sg.Text("     ENTRADA DE ITENS", font=("ANY", 18))],
    [sg.Text("")],
    [sg.Text("")],
    [sg.Text("")],

    [sg.Text("Numero", size=(15, 1),font=("Any",12)), sg.Input(size=(15, 1),font=("Any",12), key="-N-")],
    [sg.Text("Item", size=(15, 1),font=("Any",12)), sg.Input(size=(15, 1),font=("Any",12), key="-ITEM-")],
    [sg.Text("Ean", size=(15, 1),font=("Any",12)), sg.Input(size=(15, 1),font=("Any",12), key="-EAN-")],
    [sg.Text("Descrição", size=(15, 1),font=("Any",12)), sg.Input(size=(15, 1),font=("Any",12), key="-DESC-")],
    [sg.Text("Quantidade", size=(15, 1),font=("Any",12)), sg.Input(size=(15, 1),font=("Any",12), key="-QTD-")],
    [sg.Text("Preço", size=(15, 1),font=("Any",12)), sg.Input(size=(15, 1),font=("Any",12), key="-PRECO-")],
    [sg.Text("")],
    [sg.Text("")],
    [sg.Text("")],
    [sg.Button("REGISTRAR",size=(12)),sg.Button("PLANILHA",size=(12)), sg.Button("SAIR",size=(7))]
]

layout = [
    [sg.Col(entradas), sg.Col(tabela)],
]

window = sg.Window("Tabela com valores", layout, size=(810, 400), resizable=True)

while True:
    event, values = window.read()

    if event in (sg.WIN_CLOSED, "SAIR"):
        break
    elif event == "REGISTRAR":
        # Obter valores dos campos de entrada
        novo_registro = [
            values["-N-"],
            values["-ITEM-"],
            values["-EAN-"],
            values["-DESC-"],
            values["-QTD-"],
            values["-PRECO-"]
        ]

        # Adicionar os valores à lista
        lista.append(novo_registro)

        # Atualizar a tabela com os novos valores
        window["-TABELA-"].update(values=lista)


    elif event == "GERAR PLANILHA":
        df = pd.DataFrame(lista[1:], columns=titulos)
        file_path = sg.popup_get_file("Salvar Planilha com Python", save_as=True, default_extension=".xlsx")
        
        if file_path:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
                pd.save()
        sg.popup("Planilha salva")

window.close()
