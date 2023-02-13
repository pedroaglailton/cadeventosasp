<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level1.inc"-->
<%
If session("allow") = False Then Response.Redirect "../index.asp"
If session("clearance") < 1 Then Response.Redirect "utility.asp?method=unauthorized"

Const RegPorPag = 20
VarPagMax = 10
cor_linha_selecionada = "gainsboro"
pagina_consulta = "df_consulta.asp"
pagina_alteracao = "df_alteracao.asp"
pagina_inclusao = "df_inclusao.asp"
pagina_login = "../index.asp"
pagina_imprimir = "df_imprimir.asp"

If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then

%>
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
<html lang="en" prefix="og: http://ogp.me/ns#">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Sistema de Inscricao</title>
    <meta name="description" content="Learn WebSlides in simple steps starting from basic to advanced concepts with examples.">
    <link href="https://fonts.googleapis.com/css?family=Roboto:100,100i,300,300i,400,400i,700,700i%7CMaitree:200,300,400,600,700&amp;subset=latin-ext" rel="stylesheet">
    <link rel="stylesheet" type='text/css' media='all' href="static/css/base.css">
    <link rel="stylesheet" type='text/css' media='all' href="static/css/colors.css">
    <link rel="stylesheet" type='text/css' media='all' href="static/css/svg-icons.css">
    <link rel="shortcut icon" sizes="16x16" href="static/images/favicons/favicon.png">      
    <meta name="mobile-web-app-capable" content="yes">
    <meta name="theme-color" content="#333333">
  </head>
  <body>
       <section class="bg-secondary">
          <span class="background dark" style="background-image:url('https://source.unsplash.com/RkBTPqPEGDo/')"></span> 	      
          <div class="wrap">
          <fieldset>
		   <!--.wrap = container (width: 90%) -->
          
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
	
    If indice <> "" Then strQ_delete = " SELECT * FROM inc_alunos WHERE " & indice

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
 ' End If

Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.CursorLocation = 3
objRS.CursorType = 2
objRS.LockType = 1
objRS.CacheSize = RegPorPag
strQ = "SELECT * FROM inc_alunos WHERE cnpj_escola = '"&session("admin")&"'"
If Trim(Request("string_busca")) <> "" Then
  If Trim(Request("campo_busca")) <> "" Then
    strQ = strQ & " and " & Trim(Request("campo_busca")) & " LIKE '%" & Trim(Request("string_busca")) & "%'"
  Else
    strQ = strQ & " Or nome LIKE '%" & Trim(Request("string_busca")) & "%'"
	strQ = strQ & " Or responsavel LIKE '%" & Trim(Request("string_busca")) & "%'"
	strQ = strQ & " Or categorias LIKE '%" & Trim(Request("string_busca")) & "%'"
	strQ = strQ & " Or matricula_escola LIKE '%" & Trim(Request("string_busca")) & "%'"
	strQ = strQ & " Or sexo LIKE '%" & Trim(Request("string_busca")) & "%'"
    
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
strQ_indice = "SELECT * FROM inc_alunos WHERE 1 <> 1"
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

<img src="imagens/logo1.png" class="aligncenter" width="122" height="149" >
<h2 class="aligncenter">Consultar Alunos</h2>
      <p class="aligncenter">Sistema de inscrição de eventos<br>
      <a href="http://www.sistemainc.com.br/cearacolegial/sistema/welcome.asp">Voltar </a></p>
     
<BR>
Visualize os registros da 
tabela abaixo:<BR>
<FORM name="form_busca" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME") %>">
Pesquisar por <INPUT type=text name=string_busca value="<%=Request("string_busca") %>" class=texto_pagina> em
<SELECT name=campo_busca class=texto_pagina>
  <OPTION value="nome" <% If Trim(Request("campo_busca")) = Trim("nome") Then : Response.Write "selected" : End If %>>Nome</OPTION>
  <option value="rg">Identidade</option>
  <option value="cpf">cpf</option>
   <option value="data_nascimento">Data Nascimento</option>
   <option value="sexo">sexo</option>
   
  </SELECT>
<INPUT type="submit" name="submit" value="ok" class=texto_pagina style="color: black">
</FORM>

<%
If Not(objRS.EOF) Then
  objRS.AbsolutePage = PagAtual
  TotPag = objRS.PageCount
%>

Foram encontrados <%= objRS.RecordCount%> alunos cadastrados <BR>
<BR>

<TABLE >
  <TR >

<%
If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
 Response.Write "<TD align=""center"" style=""background-color: blue; color: white"" width=""1%"" nowrap><b>Editar..</b></TD>"
 End IF
 
If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
 Response.Write "<TD align=""center"" style=""background-color: blue; color: white"" width=""1%"" nowrap><b>Imprimir</b></TD>"
 End IF
 

If Right(Request.QueryString("Ordem"), 3) = "asc" Then
  Ordem = "desc"
Else
  Ordem = "asc"
End IF
%>

<TD width="4%"><b>FOTO</b></TD>
<TD width="31%"><b>NOME ALUNO</b></TD>
<TD width="8%"><b>SEXO</b></TD>
<TD width="12%"><b>NASCIMENTO</b></TD>
<TD width="15%"><b>MATRICULA</b></TD>
<TD width="17%"><b>CPF</b></TD>
<TD width="13%"><b>IDENTIDADE(RG)</b></TD>
   </TR>   
<%
For Cont = 1 to objRS.PageSize
%>

  <TR class=exibe_registros onMouseOver="this.style.backgroundColor='<%=cor_linha_selecionada%>';" onMouseOut="this.style.backgroundColor='';">

<%

If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
  Response.Write "<FORM name=""form_edit_" & Cont & """ action=""" & pagina_alteracao & """ method=post>"
  Response.Write "<TD  align=""center"" nowrap style=""background-color: ""  nowrap>&nbsp;"  
  If indice <> "" Then Response.Write "<input type=""hidden"" name=""indice"" value=""" & indice & "=" & objRS.Fields.Item(indice).Value & """>"
  Response.Write "<INPUT type=hidden name=recordno value=""" & (objRS.AbsolutePosition) & """>"
  Response.Write "<INPUT type=hidden name=strQ value=""" & strQ & """>"
  Response.Write "<INPUT type=image src=""imagens\edit_new.fw.png"" alt=""Alterar Registro"" name=alterar value=alterar>"  
  If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
  End If
  Response.Write "</TD>"
  Response.Write "</FORM>"
End If
If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
  Response.Write "<FORM name=""form_edit_" & Cont & """ action=""" & pagina_imprimir & """ method=post>"
  Response.Write "<TD  align=""center"" nowrap style=""background-color: ""  nowrap>"
  If indice <> "" Then Response.Write "<input type=""hidden"" name=""indice"" value=""" & indice & "=" & objRS.Fields.Item(indice).Value & """>"
  Response.Write "<INPUT type=hidden name=recordno value=""" & (objRS.AbsolutePosition) & """>"
  Response.Write "<INPUT type=hidden name=strQ value=""" & strQ & """>"           
  Response.Write "<INPUT  type=image src=""imagens\print_new.fw.png"" alt=""imprimir Registro"" name=imprimir value=imprimir>"
  If Session("admin") <> "" And Session("ip_admin") = Request.ServerVariables("REMOTE_ADDR") Then
  End If
  Response.Write "</TD>"
  Response.Write "</FORM>"
End If
%>
     <TD><a href="<%=(objRS.Fields.Item("foto").Value)%>" onClick="abre_foto(350, 350, 'janela_foto')" target="janela_foto"><img src="<%=(objRS.Fields.Item("foto").Value)%>" border=0 width=50 height="54"></a></TD>
	 <TD><%=(objRS.Fields.Item("nome").Value)%></TD>
	 <TD><%=(objRS.Fields.Item("sexo").Value)%></TD>
	  <TD><%=(objRS.Fields.Item("data_nascimento").Value)%></TD>
    <TD><%=(objRS.Fields.Item("matricula_escola").Value)%></TD>
	 <TD><%=(objRS.Fields.Item("cpf").Value)%></TD>
	 <TD><%=(objRS.Fields.Item("rg").Value)%></TD>
	
  </TR>

<%
  objRS.MoveNext
  If objRS.Eof then Exit For
Next
Set Cont = Nothing
%>

<TR>
  <TD colspan="18"><%LinksNavegacao()%></TD>
</TR>

</TABLE>

<%
  If indice = "" Then
    
  End If
  objRS.Close
  Set objRS = Nothing
Else
%>

<BR><B>Nenhum registro foi encontrado</B><BR><BR>

<%
End If
%>

</BODY>
</HTML>



<%
Sub LinksNavegacao()
'O código a seguir insere uma tabela com todos os links de navegação das páginas
Response.Write "<TABLE border=0 cellPadding=2 cellSpacing=0 class=tabela_paginacao>"
Response.Write "<TR><TD align=center vAlign=top noWrap colspan=5>"
Response.Write "Página " & PagAtual & " de " & TotPag
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
end if
end if

%>
           
                       </div>
          <!-- .end .wrap -->
        </section>
        
		   </section>
     <script src="../static/js/webslides.js"></script>    
    <script>
      window.ws = new WebSlides();
    </script>
    
    <!-- OPTIONAL - svg-icons.js (fontastic.me - Font Awesome as svg icons) -->
    <script defer src="../static/js/svg-icons.js"></script>
  </fieldset>
  </body>
</html>