<!DOCTYPE html>
<html>
	<head>
		<meta charset="utf-8">
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<title>Mini Mercado Teste</title>
		<style> table, td { border: 1px solid red; }</style>
	</head>
	<body>
		<%
			Dim con
			Dim rec
			Dim rs
			Dim x

			Set con=Server.createObject("Adodb.Connection")

			Set rec=Server.createObject("Adodb.recordset")

			con.open "mini-mercado"

			Set rs= con.execute("select * from cliente")
		%>
		<table>
			<tr>
				<td>Nome</td>
				<td>Email</td>
				<td>Senha</td>
			</tr>
			<%
				Do Until rs.EOF
					Response.write("<tr>")
						For Each x In rs.Fields
							Response.write("<td>" & x.value & "</td>")
						Next
					Response.write("</tr>")
					rs.movenext
				Loop

			%>
		</table>
	</body>
</html>