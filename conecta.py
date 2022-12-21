import pyodbc

dados_conexao = (
    "Driver={Devart ODBC Driver for Oracle};"
    "Server=10.1.0.137;"
    "PORT:1521;"
    "Database=DBFESOCONS;"
    "Username=rm;"
    "Password=f/M701iv_LoAE1@;"
    "SERVICE_NAME=dbfeso2;"
)

conexao = pyodbc.connect(dados_conexao)
cursor = conexao.cursor()

