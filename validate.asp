<%@Language=VBScript%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<html>
<body>
<%
cnpj = Replace(Trim(Request.Form("cnpj")), "'", "''")
senha = Replace(Trim(Request.Form("senha")), "'", "''")

If cnpj = "" OR senha = "" Then Response.Redirect "index.asp"

'SQL = "Select Registro, cnpj, senha, Clearance, ExpireDate, nome_fantasia, datacadastro, nome_escola, endereco_escola, From inc_escolas"
strQ = "SELECT * FROM inc_escolas  " 
Set objRS = MyConn.Execute(strQ)
  
 
    
	
While Not objRS.EOF  
  If cnpj = objRS("cnpj") And senha = objRS("senha") Then
    If objRS("ExpireDate") > Now() Then
      Session("allow") = True
      Session("clearance") = objRS("Clearance")
	  Session("nome_fantasia") = objRS("nome_fantasia")
	  Session("nome_escola") = objRS("nome_escola")
	  Session("datacadastro") = objRS("datacadastro")
	  Session("Registro") = objRS("Registro")
	  Session("admin") = cnpj
      Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR")
      Level = objRS("Clearance")
    Else
      Response.Redirect "utility.asp?method=expired"
    End If
  End If
  objRS.MoveNext
Wend

CleanUp(objRS)

If Session("allow") = True Then
  If Level = 3 Then Response.Redirect "admin.asp"
  'If Level < 3 Then Response.Redirect "welcome.asp"
  If Level < 3 Then Response.Redirect "welcome.asp"
Else
  Response.Redirect "index.asp"
End If

%>

</body>
</html>
