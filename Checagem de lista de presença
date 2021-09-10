import io
import openpyxl
from openpyxl import load_workbook

# arquivo OK
arquivo = open('presenca', encoding = 'utf-8').read()
arquivo = list(arquivo.splitlines())

lista = [
    'Leonardo Cardozo', 'Alex Ricardo', 'Jonathas Souza', 'Isa Catelani', 
    'Rodolfo Azevedo', 'Nathiele Cordeiro', 'Marcinhojr Correa', 'Geovani Chiapesan',
    'Bruno Margonar', 'Rafael Simao', 'Gabriela Fuzaro', 'Livia Tornai',
    'Rafael Silva', 'leonardo mustafa pastori', 'Maria  Eduarda Bigoni',
    'Ana Julia Demarque', 'Wesley Parra', 'jean augusto napedri',
    'Diego Carminati Candido', 'Fábio Belém'
    ]

# checar com os nomes da lista se estão no arquivo
print ([i.lower() for i in arquivo if i in lista])

# instanciar o arquivo
wb = load_workbook('frequencia.xlsx')
print(wb.sheetnames)

# trazer a página para um objeto
sheet = wb['Planilha1']

def loop(sheet):
    for i in range(4,38):
        select = sheet.cell(row = i, column = 2).value.lower()
        #junto = list(select)
        print(select)

plan = loop(sheet)
type(plan)

# está como None - será retomado 
print([i for i in plan if i in lista])

# Continue...
