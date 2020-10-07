<%
titulo=Request.Form("titulo")
genero=Request.Form("genero")
pag = "http://localhost/5ainfi/roberta/inicio.html"
aluga = "http://localhost/5ainfi/roberta/inicio.html"

Dim objConex, strCaminho
strCaminho = Server.MapPath("locadora.mdb")
Set objConex = Server.CreateObject("ADODB.Connection")
objConex.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strCaminho & ";"

selsql = "Select * from filmes"
set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.open selsql, objConex

dim regExiste
regExiste = false

Do While not (objRS.eof)
	if StrComp(objRS("titulo"),titulo, vbTextCompare) = 0 OR StrComp(objRS("genero"), genero, vbTextCompare) = 0 then
		Response.write("<html><HEAD><LINK REL=StyleSheet HREF=""estilos.css"" TYPE=""text/css""></HEAD><center>")
		Response.Write("<br>")
		Response.Write("<img src=")
		Response.Write(ObjRS.Fields("imagem"))
		Response.Write(" width=200 heigth=80 border=1>")
		Response.Write("<br><br> <b>Título: </b>")
		Response.write(ObjRS.Fields("titulo"))
		Response.Write("<br> <b>Gênero: </b>")
		Response.write(ObjRS.Fields("genero"))
		Response.Write("<br>")
		Response.Write("<br> <b>Direção: </b>")
		Response.write(ObjRS.Fields("direcao"))
		Response.Write("<br>")
		Response.Write("<br> <b>Sinopse: </b>")
		Response.write(ObjRS.Fields("sinopse"))
		Response.Write("<br>")
		Response.write("<a href=" & aluga & ">Alugar</a>")
		Response.write(" | ")
		Response.write("<a href=" & pag & ">Voltar</a>")
		Response.Write("</center></html>")
		regExiste = true
	end if
	
	objRS.MoveNext
Loop
	if regExiste=false then
		Response.write("<html><HEAD><LINK REL=StyleSheet HREF=""estilos.css"" TYPE=""text/css""></HEAD><center>")
		Response.Write("<br>")
		Response.write("Nenhuma resposta foi encontrada.")
		Response.Write("<br>")
		Response.write("<a href=" & pag & ">Voltar</a>")
		Response.Write("</center></html>")
	end if
	
	objRS.Close
	set objRS = Nothing
	objConex.Close
	set objConex = Nothing
%>
