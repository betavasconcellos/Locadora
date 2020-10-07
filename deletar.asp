<%

login= Request.Form("login")
pag = "http://localhost/5ainfi/roberta/index_admin.html"


Dim objConex, strCaminho,objRS
strCaminho = Server.MapPath("locadora.mdb")
Set objConex = Server.CreateObject("ADODB.Connection")
objConex.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strCaminho & ";"
Set objRS = Server.CreateObject("ADODB.Recordset")


delsql = "Delete from cadastro where login = '" & login & "'"
		
		objRS.open delsql, objConex
		set objRS = ObjConex.Execute(delsql)
		Response.write("<html><HEAD><LINK REL=StyleSheet HREF=""estilos.css"" TYPE=""text/css""></HEAD>")
		Response.Write("<br>")
		Response.write("<center><h4>Registro Excluído!</h4>")
		Response.write("<a href=" & pag & ">Voltar</a></center></html>")
		
'objRS.Close
set objRS = Nothing
objConex.Close
set objConex = Nothing

%>
