# Documentação do Software Cadastro de Itens

O **Cadastro de Itens** é um software desenvolvido em Python utilizando as bibliotecas PySimpleGUI e pandas. Ele oferece uma interface gráfica simples para o usuário inserir informações sobre itens, como nome, código EAN, descrição, quantidade e preço, e permite exportar esses dados para uma planilha eletrônica .

## Recursos Principais

- **Interface Gráfica Intuitiva**:
  - Utilização de PySimpleGUI para construção da interface gráfica.
  - Inclusão de imagem de login para uma experiência visual mais completa.

- **Funcionalidades**:
  - Inserção de informações sobre o item (nome, EAN, descrição, quantidade, preço).
  - Tratamento e validação dos dados inseridos na interface principal `try Except`.
  - Registro dos itens em uma lista primaria.
  - Visualização dos itens registrados em uma tabela.
  - Exportação dos dados para uma planilha eletrônica em formato Excel (.xlsx).
  - Registro dos dados em um BD externo.
    
## Estrutura do Código

O código está estruturado da seguinte maneira:

- **Configuração de Interface**:
  - Definição de cores e estilos para a interface gráfica usando `sg.SetOptions`.

- **Definição de Layout**:
  - Organização dos elementos visuais (campos de entrada, botões e tabela) dentro de frames e layouts.

- **Loop Principal**:
  - Utilização de um loop principal (`while True`) para capturar eventos da interface.

- **Eventos**:
  - Manipulação de eventos como registro de item (`REGISTRAR`), exportação para planilha (`PLANILHA`) e fechamento da janela (`SAIR`).

- **Interação com pandas**:
  - Uso do pandas para manipulação e exportação dos dados para uma planilha Excel.

- **Interação com MySQL**:
  - O sistema usa SQL em sua estrutura principal para armazenamento, leitura e manipulação de dados `mysql.connector.connect`.

## Uso

1. Execute o código para iniciar a aplicação.
2. Preencha os campos de entrada com as informações do item.
3. Clique em "REGISTRAR" para adicionar o item à lista.
4. Para exportar os dados para uma planilha, clique em "PLANILHA" e selecione o local de salvamento.
5. Explore outras funcionalidades, como visualizar os itens na tabela e fechar a aplicação.

## Requisitos

- Python 3.x
- PySimpleGUI
- pandas

## Desenvolvedor

Desenvolvido por Deleon Santos.

---

Este software oferece uma solução prática para o cadastro e exportação de itens em uma interface gráfica amigável. Ele pode ser utilizado como base para desenvolver sistemas mais completos de gerenciamento de itens em Python.
# Tabelas-em-Python
