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
usuario= Request.Form("usuario")
login= Request.Form("login")
xdia=CInt(dia) 
xmes=CInt(mes)
xano=Cint(ano)
xcep=CLng(cep)
xcpf=CSng(cpf)
pag = "http://localhost/5ainfi/roberta/inicio_logado.html"

Dim objConex, strCaminho
strCaminho = Server.MapPath("locadora.mdb")
Set objConex = Server.CreateObject("ADODB.Connection")
objConex.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strCaminho & ";"
Set objRS = Server.CreateObject("ADODB.RecordSet")

selsql = "Select * from cadastro"
objRS.open selsql, objConex

if StrComp(objRS("login"), login, vbTextCompare) = 0 then
objRS("nome") = nome
objRS.Update
end if

'set objRS = objConex.Execute(upsql)
		'objRS.Fields("nome")=nome
		'objRS.Fields("dia")=xdia
		'objRS.Fields("mes")=xmes
		'objRS.Fields("ano")=xano
		'objRS.Fields("ende")=ende
		'objRS.Fields("comp")=comp
		'objRS.Fields("bairro")=bairro
		'objRS.Fields("cep")=xcep
		'objRS.Fields("senha")=senha
		'objRS.Fields("email")= email
		'set objRS = objConex.Execute("Insert into cadastro (cpf,nome,dia,mes,ano,ende,comp,bairro,cep,login,senha,email) values (xcpf,'" & nome & "',xdia,xmes,xano,'" & ende & "','" & comp & "','" & bairro & "',xcep,'" & login & "','" & senha & "','" & email & "')")
Response.Write(login)
Response.Write(nome)		
Response.Write("Registro Alterado com sucesso!")
Response.Write("<br>")
Response.Write("<a href=" & pag & ">Voltar</a>")

	'objRS.Close
	set objRS = Nothing
	objConex.Close
	set objConex = Nothing

%>