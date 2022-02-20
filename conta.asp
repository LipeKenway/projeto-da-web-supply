<%
	Option Explicit

	Dim con 'conectar objeto'
	Dim rec 'registrar o objeto'
	Dim nome 'nome'
	Dim email 'email'
	Dim senha 'senha'

	'Criando uma Conexão'

	Set con=Server.createObject("Adodb.Connection")

	'Registar o Objeto'

	Set rec=Server.createObject("Adodb.recordset")

	'Abrindo Conexão'
	con.open "mini-mercado"

	'Pegando os dados do Form'
	nome = Request.form("btn-nome")
	email = Request.form("btn-email")
	nome = Request.form("btn-senha")

	'Testando o Insert'

	con.execute("insert into cliente values(" & nome & " , " & email & ", " & senha &")")

	Response.write("Data inserida com sucesso")
%>
