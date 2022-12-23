import pyodbc

dados_conexao = ('DRIVER={Devart ODBC Driver for Oracle};Host=10.1.0.137;Port=1521;SID=DBFESOCONS;UID=GTIC_078463;Password=gtic_sup23_078463;Direct=True')

conexao = pyodbc.connect(dados_conexao)
cursor = conexao.cursor()

cursor.execute("SELECT * FROM SZPACIENTE;")
tables = cursor.fetchall()
tabela = cursor.rowVerColumns['NOMEPACIENTE']
print(tabela)