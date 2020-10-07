<%

titulo = Request.Form("titulo")
imagem = Request.Form("imagem")
direcao = Request.Form("direcao")
genero = Request.Form("genero")
sinopse = Request.Form("sinopse")

pag = "http://localhost/5ainfi/roberta/inicio_admin.html"

Dim objConex, strCaminho
strCaminho = Server.MapPath("locadora.mdb")
Set objConex = Server.CreateObject("ADODB.Connection")
objConex.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strCaminho & ";"

set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.open "Select * from filmes", objConex,1,2


dim regExiste
regExiste = false

Do While not (objRS.eof)
	if StrComp(objRS("titulo"), titulo, vbTextCompare) = 0 then
		Response.write("<html><HEAD><LINK REL=StyleSheet HREF=""estilos.css"" TYPE=""text/css""></HEAD>")
		Response.Write("<br>")
		Response.write("<center><h4>Já cadastrado!</h4>")
		Response.write("<a href=" & pag & ">Voltar</a></center></html>")
		regExiste = true
	end if
	objRS.MoveNext
Loop		
	if regExiste = false then
		objRS.AddNew
		objRS.Fields("titulo")=titulo
		objRS.Fields("imagem")=imagem
		objRS.Fields("direcao")=direcao
		objRS.Fields("genero")=genero
		objRS.Fields("sinopse")=sinopse
		objRS.Update
		
		Response.write("<html><HEAD><LINK REL=StyleSheet HREF=""estilos.css"" TYPE=""text/css""></HEAD>")
		Response.write("<br>")
		Response.Write("<center><h4>Inserido!</h4>")
		Response.write("<a href=" & pag & ">Voltar</a><center></html>")
	
	End if
	
	objRS.Close
	set objRS = Nothing
	objConex.Close
	set objConex = Nothing
%>
