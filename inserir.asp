<%

cpf = Request.Form("cpf")
nome = Request.Form("nome")
dia = Request.Form("dia")
mes = Request.Form("mes")
ano = Request.Form("ano")
ende= Request.Form("ende")
comp= Request.Form("comp")
bairro= Request.Form("bairro")
cep= Request.Form("cep")
login= Request.Form("login")
senha= Request.Form("senha")
email= Request.Form("email")
xdia=CInt(dia) 
xmes=CInt(mes)
xano=Cint(ano)
xcep=CLng(cep)
xcpf=CSng(cpf)
pag = "http://localhost/5ainfi/roberta/inicio.html"

Dim objConex, strCaminho
strCaminho = Server.MapPath("locadora.mdb")
Set objConex = Server.CreateObject("ADODB.Connection")
objConex.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strCaminho & ";"


'selsql = "Select * from cadastro"
set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.open "Select * from cadastro", objConex,1,2


dim regExiste
regExiste = false

Do While not (objRS.eof)
	if StrComp(objRS("cpf"), cpf, vbTextCompare) = 0 then
		Response.write("<html><HEAD><LINK REL=StyleSheet HREF=""estilos.css"" TYPE=""text/css""></HEAD>")
		Response.Write("<br>")
		Response.write("<center><h4>Cliente já cadastrado!</h4>")
		Response.write("<a href=" & pag & ">Voltar</a></center></html>")
		regExiste = true
	end if
	objRS.MoveNext
Loop		
	if regExiste = false then
		objRS.AddNew
		objRS.Fields("cpf")=xcpf
		objRS.Fields("nome")=nome
		objRS.Fields("dia")=xdia
		objRS.Fields("mes")=xmes
		objRS.Fields("ano")=xano
		objRS.Fields("ende")=ende
		objRS.Fields("comp")=comp
		objRS.Fields("bairro")=bairro
		objRS.Fields("cep")=xcep
		objRS.Fields("login")=login
		objRS.Fields("senha")=senha
		objRS.Fields("email")= email
		'set objRS = objConex.Execute("Insert into cadastro (cpf,nome,dia,mes,ano,ende,comp,bairro,cep,login,senha,email) values (xcpf,'" & nome & "',xdia,xmes,xano,'" & ende & "','" & comp & "','" & bairro & "',xcep,'" & login & "','" & senha & "','" & email & "')")
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
