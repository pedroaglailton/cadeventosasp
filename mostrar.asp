
<!--#INCLUDE FILE="level1.inc"-->
	
<%

strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\sistemainc\Dados\ipu\colegial.mdb"

'N?mero total de registros a serem exibidos por p?gina
Const RegPorPag =1

'N?mero de p?ginas a ser exibido no ?ndice de pagina??o
VarPagMax = 1

%>
<style type="text/css">
<!--
.texto_pagina
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
color: dimgray;
}

.tabela_registros
{
width: 100%;
background-color: white;
}

.titulos_registros
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
color: white;
background-color: #6699CC;
}

.exibe_registros
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
width: 100%;
color: dimgray;
background-color: aliceblue;
}

.tabela_paginacao
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
width: 100%;
color: gray;
border-top: 1px solid gainsboro;
background-color: whitesmoke;
}

.links_paginacao
{
color: royalblue;
text-decoration: none;
}

.links_paginacao:hover
{
color: gray;
text-decoration: underline;
}
-->
</style>
<SCRIPT language="JavaScript">
<!--
function abre_foto(width, height, nome) {
  var top; var left;
  top = ( (screen.height/2) - (height/2) )
  left = ( (screen.width/2) - (width/2) )
  window.open('',nome,'width='+width+',height='+height+',scrollbars=yes,toolbar=no,location=no,status=no,menubar=no,resizable=no,left='+left+',top='+top);
}
function confirm_delete(form) {
  if (confirm("Tem certeza que deseja excluir o registro?")) {
	document[form].action = '<%=Request.ServerVariables("SCRIPT_NAME")%>';
	document[form].submit();
  }
}
//-->
</SCRIPT>
<SCRIPT language="JavaScript">
Function HTMLEspeciais(sString)
 If (sString <> "") Then
  sString = Replace(sString, "á", "&aacute;")
  sString = Replace(sString, "â", "&acirc;")
  sString = Replace(sString, "à", "&agrave;")
  sString = Replace(sString, "ã", "&atilde;")

  sString = Replace(sString, "ç", "&ccedil;")

  sString = Replace(sString, "é", "&eacute;")
  sString = Replace(sString, "ê", "&ecirc;")

  sString = Replace(sString, "í", "&iacute;")

  sString = Replace(sString, "ó", "&oacute;")
  sString = Replace(sString, "ô", "&ocirc;")
  sString = Replace(sString, "õ", "&otilde;")

  sString = Replace(sString, "ú", "&uacute;")
  sString = Replace(sString, "ü", "&uuml;")

  sString = Replace(sString, "Á", "&Aacute;")
  sString = Replace(sString, "Â", "&Acirc;")
  sString = Replace(sString, "À", "&Agrave;")
  sString = Replace(sString, "Ã", "&Atilde;")

  sString = Replace(sString, "Ç", "&Ccedil;")

  sString = Replace(sString, "É", "&Eacute;")
  sString = Replace(sString, "Ê", "&Ecirc;")

  sString = Replace(sString, "Í", "&Iacute;")

  sString = Replace(sString, "Ó", "&Oacute;")
  sString = Replace(sString, "Ô", "&Ocirc;")
  sString = Replace(sString, "Õ", "&Otilde;")

  sString = Replace(sString, "Ú", "&Uacute;")
  sString = Replace(sString, "Ü", "&Uuml;")

  sString = Replace(sString, """", "&quot;") 
  sString = Replace(sString, "<", "&lt;") 
  sString = Replace(sString, ">", "&gt;") 
 End If

 HTMLEspeciais = sString
End Function

</SCRIPT>
</HEAD>
<BODY class=texto_pagina>
<%
If Request.QueryString("PagAtual") = "" Then
  PagAtual = 1
  NumPagMax = VarPagMax
Else
  NumPagMax = CInt(Request.QueryString("NumPagMax"))
  PagAtual = CInt(Request.QueryString("PagAtual"))
  Select Case Request.QueryString("Submit")
    Case "Anterior" : PagAtual = PagAtual - 1
    Case "Proxima" : PagAtual = PagAtual + 1
    Case "Menos" : NumPagMax = NumPagMax - VarPagMax
    Case "Mais" : NumPagMax = NumPagMax + VarPagMax
    Case Else : PagAtual = CInt(Request.QueryString("Submit"))
  End Select
  If NumPagMax < PagAtual then
    NumPagMax = NumPagMax + VarPagMax
  End If
  If NumPagMax - (VarPagMax - 1) > PagAtual then
    NumPagMax = NumPagMax - VarPagMax
  End If
End If

Set objCon = Server.CreateObject("ADODB.Connection")
objCon.Open strCon

  If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
  If Request.Form("recordno") <> "" Then
    Set objRS_delete = Server.CreateObject("ADODB.Recordset")
    objRS_delete.CursorLocation = 3
    objRS_delete.CursorType = 0
    objRS_delete.LockType = 3

    strQ_delete = Request.Form("strQ")
    indice = Trim(Request.Form("indice"))
    If indice <> "" Then strQ_delete = " SELECT * FROM informes WHERE " & indice

    objRS_delete.Open strQ_delete, objCon, , , &H0001
    If indice = "" Then objRS_delete.Move Request.Form("recordno") - 1
    If Not objRS_delete.EOF Then
      objRS_delete.Delete
      objRS_delete.UpdateBatch
    End IF

    objRS_delete.Close
    Set objRS_delete = Nothing
    Set strQ_delete = Nothing
  End If
  End If

Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.CursorLocation = 3
objRS.CursorType = 2
objRS.LockType = 1
objRS.CacheSize = RegPorPag
strQ = "SELECT * FROM informes"

If Trim(Request("string_busca")) <> "" Then
  If Trim(Request("campo_busca")) <> "" Then
    strQ = strQ & " Where " & Trim(Request("campo_busca")) & " LIKE '%" & Trim(Request("string_busca")) & "%'"
  Else
    strQ = strQ & " Where 1 <> 1"
    strQ = strQ & " Or informes LIKE '%" & Trim(Request("string_busca")) & "%'"
  End If
End If

If Trim(Request.QueryString("Ordem")) <> "" Then
  strQ = strQ & " ORDER BY " & Request.QueryString("Ordem")
End If
objRS.Open strQ, objCon, , , &H0001
objRS.PageSize = RegPorPag

Set objRS_indice = Server.CreateObject("ADODB.Recordset")
objRS_indice.CursorLocation = 2
objRS_indice.CursorType = 0
objRS_indice.LockType = 2
strQ_indice = "SELECT * FROM informes WHERE 1 <> 1"
objRS_indice.Open strQ_indice, objCon, , , &H0001
indice = ""
For Each item In objRS_indice.Fields
  If item.properties("IsAutoIncrement") = True Then
    indice = item.name
    Exit For
  End If
Next
objRS_indice.Close
Set objRS_indice = Nothing
Set strQ_indice = Nothing

Set objRS.ActiveConnection = Nothing
objCon.Close
Set objCon = Nothing
%>
<BR>
<%
If Not(objRS.EOF) Then
  objRS.AbsolutePage = PagAtual
  TotPag = objRS.PageCount
%>

Foram encontrados <%= objRS.RecordCount%> Informe(s)<BR>
<BR>

<TABLE border=0 cellpadding=2 cellspacing=1 class=tabela_registros>
  <TR class=titulos_registros>

<%

If Right(Request.QueryString("Ordem"), 3) = "asc" Then
  Ordem = "desc"
Else
  Ordem = "asc"
End IF
%>

  
<%
For Cont = 1 to objRS.PageSize
%>

  <TR class=exibe_registros onMouseOver="this.style.backgroundColor='<%=cor_linha_selecionada%>';" onMouseOut="this.style.backgroundColor='';">

<%

  If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
   End If
 
%>

    <textarea name="textarea" readonly cols="" rows="" ><%=(objRS.Fields.Item("informes").Value)%> </textarea>
  

<%
  objRS.MoveNext
  If objRS.Eof then Exit For
Next
Set Cont = Nothing
%>

<TR>
  <TD colspan="3"><%LinksNavegacao()%></TD>
</TR>

</TABLE>


<BR>

<%
End If
%>


</BODY>
</html>



<%
Sub LinksNavegacao()
'O c?digo a seguir insere uma tabela com todos os links de navega??o das p?ginas
Response.Write "<TABLE border=0 cellPadding=2 cellSpacing=0 class=tabela_paginacao>"
Response.Write "<TR><TD align=center vAlign=top noWrap colspan=5>"
Response.Write "Informes " & PagAtual & " de " & TotPag
Response.Write "</TD></TR><TR><TD width=33% align=right vAlign=top noWrap>"
If PagAtual > 1 Then
  Response.Write "<A href=""" & Request.ServerVariables("SCRIPT_NAME") & "?PagAtual=" &  PagAtual &"&VarPagMax=" & VarPagMax & "&NumPagMax=" & NumPagMax & "&Submit=Anterior&Ordem=" & Request.QueryString("Ordem")& "&string_busca=" & Server.URLEncode(Request("string_busca")) & "&campo_busca=" & Server.URLEncode(Request("campo_busca"))  & """ class=links_paginacao>&lt; Anterior</A>"
End If
Response.Write "</TD><TD width=33% align=middle vAlign=top noWrap>"
If NumPagMax - VarPagMax <> 0 then
  Response.Write "&nbsp;<A href=""" & Request.ServerVariables("SCRIPT_NAME") & "?PagAtual=" & NumPagMax - VarPagMax & "&VarPagMax=" & VarPagMax & "&NumPagMax=" & NumPagMax - VarPagMax & "&Submit=Menos&Ordem=" & Request.QueryString("Ordem") & "&string_busca=" & Server.URLEncode(Request("string_busca")) & "&campo_busca=" & Server.URLEncode(Request("campo_busca")) & """ class=links_paginacao>&lt;&lt;</A>&nbsp;&nbsp;"
End If
for i = NumPagMax - (VarPagMax - 1) to NumPagMax
  If i <= TotPag then
    If i <> CInt(PagAtual) then
      Response.Write "&nbsp;<A href=""" & Request.ServerVariables("SCRIPT_NAME") & "?PagAtual=" & PagAtual & "&VarPagMax=" & VarPagMax & "&NumPagMax=" & NumPagMax & "&Submit=" & i & "&Ordem=" & Request.QueryString("Ordem") & "&string_busca=" & Server.URLEncode(Request("string_busca")) & "&campo_busca=" & Server.URLEncode(Request("campo_busca")) & """ class=links_paginacao>" & i & "</A>&nbsp;"
    Else
      If PagAtual <> TotPag Then
        Response.Write "&nbsp;" & i & "&nbsp;"
      End If
    End If
  End If
Next
If NumPagMax  < TotPag then
  Response.Write "&nbsp;&nbsp;<A href=""" & Request.ServerVariables("SCRIPT_NAME") & "?PagAtual=" & NumPagMax + 1 & "&VarPagMax=" & VarPagMax & "&NumPagMax=" & NumPagMax + VarPagMax & "&Submit=Mais&Ordem=" & Request.QueryString("Ordem") & "&string_busca=" & Server.URLEncode(Request("string_busca")) & "&campo_busca=" & Server.URLEncode(Request("campo_busca")) & """ class=links_paginacao>&gt;&gt;</A>"
End If
Response.Write "</TD><TD width=33% align=left vAlign=top noWrap>"
If PagAtual <> TotPag Then
  Response.Write "&nbsp;&nbsp;<A href=""" & Request.ServerVariables("SCRIPT_NAME") & "?PagAtual=" & PagAtual & "&VarPagMax=" & VarPagMax & "&NumPagMax=" & NumPagMax & "&Submit=Proxima&Ordem=" & Request.QueryString("Ordem") & "&string_busca=" & Server.URLEncode(Request("string_busca")) & "&campo_busca=" & Server.URLEncode(Request("campo_busca")) & """ class=links_paginacao>Proxima &gt;</A>"
End If
Response.Write "</TD></TR></TABLE>"
End Sub
%>
