<%
Dim objConex, strCaminho
strCaminho = Server.MapPath("locadora.mdb")
Set objConex = Server.CreateObject("ADODB.Connection")
objConex.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strCaminho & ";"

selsql = "Select * from cadastro where nome = nome"
nome = Request.Form("nome")
Response.write(nome)
%>