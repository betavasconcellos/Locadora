<%
login = Request.Form("login")
senha= Request.Form("senha")
pag = "http://localhost/5ainfi/roberta/inicio.html"
pag_log = "http://localhost/5ainfi/roberta/inicio_logado.html"

Dim objConex, strCaminho
strCaminho = Server.MapPath("locadora.mdb")
Set objConex = Server.CreateObject("ADODB.Connection")
objConex.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strCaminho & ";"

selsql = "Select * from cadastro"
set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.open selsql, objConex

dim regExiste
regExiste = false

Do While not (objRS.eof)
	
	'if objRS("login")= admin AND StrComp(objRS("senha"), senha, vbTextCompare) = 0 then
	'Response.Redirect ("http://localhost/5ainfi/roberta/inicio_admin.html")
	'end if
	'objRS.MoveNext
	
	if StrComp(objRS("login"), login, vbTextCompare) = 0 AND StrComp(objRS("senha"), senha, vbTextCompare) = 0 then
		if objRS("login")= "admin" AND StrComp(objRS("senha"), senha, vbTextCompare) = 0 then
		Response.Redirect ("http://localhost/5ainfi/roberta/inicio_admin.html")
		objRS.MoveNext
		else
		Response.Redirect ("http://localhost/5ainfi/roberta/inicio_logado.html")	
		end if
	end if
	objRS.MoveNext

Loop	
	if regExiste=false then
		Response.write("<html><HEAD><LINK REL=StyleSheet HREF=""estilos.css"" TYPE=""text/css""></HEAD>")
		Response.Write("<br>")
		Response.write("<center><h4>Usuário inexistente!</h4>")
		Response.write("<a href=" & pag & ">Voltar</a></center></html>")
		regExiste = true
	end if
	
	objRS.Close
	set objRS = Nothing
	objConex.Close
	set objConex = Nothing
%>
