<%@Language=VBScript%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level3.inc"-->

<html>
<body>
<center>
  <table border="0" bgcolor="#336699">
    <form action="update.asp?method=Add" method="Post">
      <tr> 
        <td><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>cnpj</b></font></td>
        <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
          <input type="text" name="cnpj" size="10">
          </font></td>
      </tr>
      <tr> 
        <td><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>Password</b></font></td>
        <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
          <input type="password" name="senha" size="10">
          </font></td>
      </tr>
      <tr> 
        <td><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>Nivel</b> 
          (1 - 3)</font></td>
        <td> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
          <select name="level">
            <option value="1">1 
            <option value="2">2 
            <option value="3">Admin 
          </select>
          </font></td>
      </tr>
      <tr> 
        <td><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>Data 
          de Expira&ccedil;&atilde;o</b></font></td>
        <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
          <input type="text" name="expdate" size="10" value="<%=DateAdd("yyyy", 1, Date)%>">
          </font></td>
      </tr>
      <tr> 
        <td bgcolor="#336699"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
          <input type="submit" value="Adicionar">
          </font></td>
        <td bgcolor="#336699">&nbsp;</tr>
    </form>
  </table>
</center>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
<%
SQL = "Select Registro, cnpj, senha, Clearance, ExpireDate From inc_escolas Order By Registro"
Set objRS = MyConn.Execute(SQL)

Response.Write "<center>"

While Not objRS.EOF
  Response.Write "<form name=""Update"" method=""Post"">"
  Response.Write "<table border=""0"" bgcolor=""#336699"">"

  %>
</font> 
<tr> 
  <td><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">CNPJ</font></td>
  <td><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Senha</font></td>
  <td><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Nivel  Data de Expiração</font></td>
</tr>
<tr> </tr>
<tr> 
  <td><input type="hidden" name="id" value="<%=objRS("Registro")%>"></td>
</tr>
<tr> 
  <td><input type="text" name="cnpj" size="10" value="<%=objRS("cnpj")%>"></td>
  <td><input type="text" name="senha" size="10" value="<%=objRS("senha")%>"></td>
  <td><input type="text" name="level" size="1" value="<%=objRS("Clearance")%>">&nbsp;&nbsp;<input type="text" name="expdate" size="10" value="<%=objRS("ExpireDate")%>"></td>
  <td bgcolor="#336699"><input type="submit" value="Atualizar" onClick="this.form.action='update.asp?method=Edit';"></td>
  <td bgcolor="#336699"><input type="submit" value="Deletar" onClick="this.form.action='update.asp?method=Delete';"></td>
</tr>
<%
  Response.Write "</table>"
  Response.Write "</form>"
  objRS.MoveNext
Wend

Response.Write "</center>"    

CleanUp(objRS)

Response.Write "<p><center><a href=""utility.asp?method=abandon""><b>Sair</b></a></center>"
%>
</font> 
</body>
</html>
