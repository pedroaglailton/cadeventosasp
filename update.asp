<%@Language=VBScript%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level3.inc"-->

<%
Dim Method
Method = Request.QueryString("method")
Select Case Method
  Case "Add"
    Add(MyConn)
  Case "Edit"
    Edit(MyConn)
  Case "Delete"
    Delete(MyConn)
End Select

'/////////////////////////////////////////////////////////////////////////////////

Sub Add(MyConn)
  Dim cnpj, senha, Level, ExpDate, SQL, Registro
  UserName = Replace(Trim(Request.Form("cnpj")), "'", "''")
  PassWord = Replace(Trim(Request.Form("senha")), "'", "''")
  Level = Trim(Request.Form("level"))
  ExpDate = Trim(Request.Form("expdate"))
  If cnpj = "" Or senha = "" Or Level = "" Or ExpDate = "" Then Response.Redirect "admin.asp"
  SQL = "Insert Into inc_escolas (cnpj, [senha], Clearance, ExpireDate) Values('"&cnpj&"', '"&senha&"', '"&Level&"', '"&ExpDate&"')"
  MyConn.Execute(SQL)
  CleanUp2()  
  Response.Redirect "admin.asp"
End Sub

'////////////////////////////////////////////////////////////////////////////////////

Sub Edit(MyConn)  
  Dim Registro, cnpj, senha, level, expdate ,SQL ,objRS
  Registro = CInt(Request.Form("Registro"))
  cnpj = Replace(Request.Form("cnpj"), "'", "''")
  senha = Replace(Request.Form("senha"), "'", "''")
  level = CInt(Request.Form("level"))
  expdate = Request.Form("expdate")
  SQL = "Update inc_escolas Set cnpj = '"&cnpj&"', [senha] = '"&senha&"'"
  SQL = SQL & ", Clearance = "&level&", ExpireDate = '"&expdate&"' Where Registro = "&Registro&""
  Set objRS = MyConn.Execute(SQL)
  CleanUp2()
  Response.Redirect "admin.asp"
End Sub

'////////////////////////////////////////////////////////////////////////////////////

Sub Delete(MyConn)
  Dim Registro, SQL
  Registro = CInt(Request.Form("Registro"))
  SQL = "Delete * From inc_escolas Where Registro = "&Registro&""
  MyConn.Execute(SQL)
  CleanUp2()  
  Response.Redirect "admin.asp"
End Sub
%>
