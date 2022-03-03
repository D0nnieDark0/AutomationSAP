from Projetos.utils.SapNew import Sap
from datetime import datetime

print(datetime.now())

# * Primeira parte do codigo Acessando o SAP e baixando os arquivos

USUARIO = ""
SENHA = ""
NUMERO_TRANSACAO = ""

s = Sap()

s.login(USUARIO, SENHA)

s.transaction(NUMERO_TRANSACAO)

s.purchase()

s.spreadsheet_account()

s.spreadsheet()

# * Segunda parte, começando a fusão dos excels

print(datetime.now())
