#Estou desenvolvendo um sistema integrado ao bd SQL que colete dados e transforme numa planilha eletronica
import PySimpleGUI as sg
import pandas as pd
import mysql.connector 
lista = []

#Definição do thema para a janela
sg.SetOptions(
                background_color='#363636', 
                text_element_background_color='#363636',
                element_background_color='#363636', 
                scrollbar_color=None, input_elements_background_color='#F7F3EC', 
                button_color=('white', '#4F4F4F'))#Configuração de thema da janela

#Definição dp layout da pagina 
titulos = ["Item", "    EAN   ", " QTD ", " Preço ","Descrição do Produto"]

imagem = [[sg.Image(filename="imagem_login.png")],]

entradas = [
    [sg.P(),sg.Frame("",imagem),sg.P() ],
    [sg.P(),sg.Text("Item", size=(6, 1),font=("Any",12)),sg.Text("Ean", size=(15, 1),font=("Any",12)),sg.Text("Qtd", size=(6, 1),font=("Any",12)),
     sg.Text("Preço", size=(15, 1),font=("Any",12)),sg.P()],
    [sg.P(),sg.Input("001",size=(6, 1),font=("Any",12), key="-ITEM-"), sg.Input(default_text="7893000252536",size=(15, 1),font=("Any",12), key="-EAN-"),
     sg.Input("12",size=(6, 1),font=("Any",12), key="-QTD-"),sg.Input("11.90",size=(15, 1),font=("Any",12), key="-PRECO-"),sg.P()],
    [sg.Text("Descrição", size=(45, 1),font=("Any",12)), sg.P()],
    [sg.P(),sg.Input("Margarina qualy 500g",size=(47, 1),font=("Any",12), key="-DESC-"),sg.P() ],
    [sg.Text("")],
    [sg.P(),sg.Button("REGISTRAR",size=(12)),sg.Button("PESQUISAR",size=(12)),sg.Button("PLANILHA",size=(12)), sg.Button("SAIR",size=(7)),sg.P()]]

layout = [
    [sg.P(),sg.Text("ENTRADA DE ITENS", font=("ANY", 25,"bold")),sg.P()],
    [sg.P(),sg.Frame("",entradas),sg.P()],
    [sg.P(),sg.Table(values=lista, headings=titulos, max_col_width=12, auto_size_columns=True,
     display_row_numbers=False, justification="center", num_rows=20, key="-TABELA-", row_height=15),sg.P()],
]
janela=[[sg.Frame("",layout)],]

window = sg.Window("Tabela com valores", janela, size=(700, 670))

#conexaão com o banco de dados MySQL
db_connection = mysql.connector.connect(
    host="localhost",
    user="root",
    password="root",
    database="cadastro_Produto")
    
cursor = db_connection.cursor()
    
while True:
    event, values = window.read()
    
    
    if event in (sg.WIN_CLOSED, "SAIR"):
        break
        
    elif event == "REGISTRAR": # Obter valores dos campos de entrada
        try:
            item=int(values["-ITEM-"].strip())
            ean=int(values["-EAN-"].strip())
            qtd=int(values["-QTD-"].strip())
            valor=float(values["-PRECO-"].strip())
            descricao=str(values["-DESC-"].strip().title())
        except ValueError:
            sg.popup('Erro em valores. ')
            continue

        if not item or not ean or not qtd or not valor or not descricao:#validação dos valores
            sg.popup('Deve preencher todos os campos')
        else:
            novo_registro=[item,ean,qtd,valor,descricao]
            lista.append(novo_registro)
            window["-TABELA-"].update(values=lista)

            insert_query = "INSERT INTO novo_item (itemId, ean,  qtd, valor, descricao) VALUES (%s, %s, %s, %s, %s)"
            cursor.execute(insert_query, novo_registro)
            db_connection.commit()

            select_query = "SELECT itemId, ean,  qtd, valor, descricao FROM novo_item"
            cursor.execute(select_query)
            retrieved_data = cursor.fetchall()

            lista = [list(row) for row in retrieved_data]
            
            window["-TABELA-"].update(values=lista)

    elif event =="PESQUISAR":
        select_query = "SELECT itemId, ean, qtd, valor, descricao FROM novo_item ORDER BY itemId ASC"
        cursor.execute(select_query)
        retrieved_data = cursor.fetchall()
        
        lista = [list(row) for row in retrieved_data]
        window["-TABELA-"].update(values=lista)

    elif event == "PLANILHA":#definição da tabela em excel
        tabela = pd.DataFrame(lista[0:], columns=titulos)
        documento = sg.popup_get_file("novo_item", save_as=True, default_extension=".xlsx")
        
        if documento:
            with pd.ExcelWriter(documento, engine='openpyxl') as escreve:
                tabela.to_excel(escreve, index=False)        
        sg.popup("Planilha salva")


cursor.close()
db_connection.close()

window.close()