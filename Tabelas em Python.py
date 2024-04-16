#Estou desenvolvendo um sistema que colete dados e transforme numa planilha eletronica
import PySimpleGUI as sg
import pandas as pd

lista = []

sg.SetOptions(
                background_color='#363636', 
                text_element_background_color='#363636',
                element_background_color='#363636', 
                scrollbar_color=None, input_elements_background_color='#F7F3EC', 
                button_color=('white', '#4F4F4F'))#Configuração de thema da janela

titulos = ["Item", "      EAN    ", "    Descrição do Produto    ", " QTD ", " Preço "]

imagem = [[sg.Image(filename="imagem_login.png")],]

entradas = [
    [sg.P(),sg.Frame("",imagem),sg.P() ],
    [sg.P(),sg.Text("Item", size=(6, 1),font=("Any",12)),sg.Text("Ean", size=(15, 1),font=("Any",12)),sg.Text("Qtd", size=(6, 1),font=("Any",12)),
     sg.Text("Preço", size=(15, 1),font=("Any",12)),sg.P()],
    [sg.P(),sg.Input("##",size=(6, 1),font=("Any",12), key="-ITEM-"), sg.Input(default_text="##",size=(15, 1),font=("Any",12), key="-EAN-"),
     sg.Input("##",size=(6, 1),font=("Any",12), key="-QTD-"),sg.Input("##",size=(15, 1),font=("Any",12), key="-PRECO-"),sg.P()],
    [sg.Text("Descrição", size=(45, 1),font=("Any",12)), sg.P()],
    [sg.P(),sg.Input("##",size=(47, 1),font=("Any",12), key="-DESC-"),sg.P() ],
    [sg.Text("")],
    [sg.P(),sg.Button("REGISTRAR",size=(12)),sg.Button("PLANILHA",size=(12)), sg.Button("SAIR",size=(7)),sg.P()]]

layout = [
    [sg.P(),sg.Text("ENTRADA DE ITENS", font=("ANY", 25,"bold")),sg.P()],
    [sg.P(),sg.Frame("",entradas),sg.P()],
    [sg.P(),sg.Table(values=lista, headings=titulos, max_col_width=12, auto_size_columns=True,
     display_row_numbers=False, justification="right", num_rows=20, key="-TABELA-", row_height=15),sg.P()],
]
janela=[[sg.Frame("",layout)],]
window = sg.Window("Tabela com valores", janela, size=(700, 670))

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, "SAIR"):
        break
        
    elif event == "REGISTRAR": # Obter valores dos campos de entrada
        novo_registro = [
            values["-ITEM-"],
            values["-EAN-"],
            values["-DESC-"],
            values["-QTD-"],
            values["-PRECO-"]]

        lista.append(novo_registro)
        window["-TABELA-"].update(values=lista)
   
    elif event == "PLANILHA":
        tabela = pd.DataFrame(lista[0:], columns=titulos)
        documento = sg.popup_get_file("Tabela_Itens", save_as=True, default_extension=".xlsx")
        
        if documento:
            with pd.ExcelWriter(documento, engine='openpyxl') as escreve:
                tabela.to_excel(escreve, index=False)        
        sg.popup("Planilha salva")

window.close()