<%
'///////////////////////////////////////////////////////////////////////////////////
'connection string


'Set MyConn=Server.CreateObject("ADODB.Connection")
'objCon = "DBQ=C:\wamp\www\CNPJ\Cadastro_Escolas\DB\colegial.mdb;Driver={Microsoft Access Driver (*.mdb)};"
'MyConn.Open objCon

Set MyConn = server.createobject("adodb.connection")
objCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\sistemainc\Dados\ipu\colegial.mdb"
MyConn.Open objCon

'///////////////////////////////////////////////////////////////////////////////////


'///////////////////////////////////////////////////////////////////////////////////
'cleanup routines

  'objCon.Close
 ' Set objCon = Nothing

Sub CleanUp(RS)
  objRS.Close
  MyConn.Close
  Set objRS = Nothing
  Set MyConn = Nothing
End Sub

Sub CleanUp2()
  MyConn.Close
  Set MyConn = Nothing
End Sub

'////////////////////////////////////////////////////////////////////////////////////
%>
