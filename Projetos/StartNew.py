from Projetos.utils.SapNew import Sap
from datetime import datetime

print(datetime.now())

# * Primeira parte do codigo Acessando o SAP e baixando os arquivos

s = Sap()

s.login("GUISOUSA", "A2U2!NKfhH8n#$w")

s.transaction("me5a")

s.purchase()

s.spreadsheet_account()

s.spreadsheet()

# * Segunda parte, começando a fusão dos excels

print(datetime.now())