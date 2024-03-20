'''
Passos:
- Ler dados da planilha.
- Inserir cada célula de cada linha eum um campo do sistema. 

Bibliotecas:
Openpyxl: Ler planilhas excel.
Pyautogui: Automação de mouse e teclado.

Terminal:
python -m venv guigadbay: Criri o ambiente virtual chamando guigadbay.
./guigadbay/Scripts/activate: Ativei o ambiente virtual.
pip install openpyxl pyautogui: Instalei as bibliotecas.
'''
import openpyxl #Permite ler planilhas.

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx') #Atribui a variavel workbook a função de abrir o arquivo venda_de_produtos.
vendas_sheet = workbook['vendas'] #Atribui a variavel vendas_sheet a função de ler a página vendas do arquivo venda_de_produtos.

#iter_rows função do Openpyxl que permite ler cada linha da planilha seguindo os parametros dos parentesis.
#min_row diz que é para começar a ler no minimo a partir da linha 2.
#linha[0].value definindo que é para a variavel linha ler cada coluna da planilha separadamente, sendo a primeira a coluna 0.
for linha in vendas_sheet.iter_rows(min_row=2):
    linha[0].value
    linha[1].value
    linha[2].value
    linha[3].value
















