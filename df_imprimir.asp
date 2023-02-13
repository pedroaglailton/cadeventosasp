<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level1.inc"-->
<%
If session("allow") = False Then Response.Redirect "../index.asp"
If session("clearance") < 1 Then Response.Redirect "utility.asp?method=unauthorized"

%>

<HTML>
<HEAD>
<TITLE>Alterar Registro</TITLE>
<meta charset="UTF-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">

<style type="text/css">
<!--
body, html { padding:0; margin:0 }
.Estilo16 {
	font-family: "Courier New", Courier, monospace;
	text-align: justify;
}
.Estilo18 {color: #666666; font-family: "Courier New", Courier, monospace; }
.Estilo19 {color: #FFFFFF}
.Estilo20 {font-size: 12px}
.Estilo23 {font-family: Arial, Helvetica, sans-serif}
.Estilo27 {color: #666666; font-family: "Courier New", Courier, monospace; font-size: 14px; }
.Estilo30 {font-size: 16px}
.Estilo5 {color: #000000}
.style7 {	color: #000000;
	font-size: 20px;
	font-family: "Ander Hedge";
	font-weight: bold;
}
.texto_pagina table tr th p {
	text-align: center;
}
-->
</style>
</HEAD>
<BODY class=texto_pagina>


<p>
  <%
If Not IsEmpty(Request.Form("salvar")) Then
  Set objCon = Server.CreateObject("ADODB.Connection")
  objCon.Open strCon
  Set objRS = Server.CreateObject("ADODB.Recordset")
  objRS.CursorLocation = 3
  objRS.CursorType = 0
  objRS.LockType = 3

  strQ = Request.Form("strQ")
  indice = Trim(Request.Form("indice"))
  If indice <> "" Then strQ = " SELECT * FROM inc_alunos WHERE " & indice

  objRS.Open strQ, objCon, , , &H0001
  If indice = "" Then objRS.Move Request.Form("recordno") - 1

  If objRS.Fields("bairro").properties("IsAutoIncrement") = False Then
    objRS("bairro") = Trim(Request.Form("bairro"))
  End If
  If objRS.Fields("cep").properties("IsAutoIncrement") = False Then
    objRS("cep") = Trim(Request.Form("cep"))
  End If
  If objRS.Fields("cidade").properties("IsAutoIncrement") = False Then
    objRS("cidade") = Trim(Request.Form("cidade"))
  End If
  If objRS.Fields("cnpj_escola").properties("IsAutoIncrement") = False Then
    objRS("cnpj_escola") = Trim(Request.Form("cnpj_escola"))
  End If
  If objRS.Fields("cpf").properties("IsAutoIncrement") = False Then
    objRS("cpf") = Trim(Request.Form("cpf"))
  End If
  If objRS.Fields("data_nascimento").properties("IsAutoIncrement") = False Then
    objRS("data_nascimento") = Trim(Request.Form("data_nascimento"))
  End If
  If objRS.Fields("email").properties("IsAutoIncrement") = False Then
    objRS("email") = Trim(Request.Form("email"))
  End If
  If objRS.Fields("endereco").properties("IsAutoIncrement") = False Then
    objRS("endereco") = Trim(Request.Form("endereco"))
  End If
  If objRS.Fields("escola").properties("IsAutoIncrement") = False Then
    objRS("escola") = Trim(Request.Form("escola"))
  End If
  If objRS.Fields("foto").properties("IsAutoIncrement") = False Then
    objRS("foto") = Trim(Request.Form("foto"))
  End If
  If objRS.Fields("Id_alunos").properties("IsAutoIncrement") = False Then
    objRS("Id_alunos") = Trim(Request.Form("Id_alunos"))
  End If
  If objRS.Fields("id_escola").properties("IsAutoIncrement") = False Then
    objRS("id_escola") = Trim(Request.Form("id_escola"))
  End If
  If objRS.Fields("modalidade1").properties("IsAutoIncrement") = False Then
    objRS("modalidade1") = Trim(Request.Form("modalidade1"))
  End If
  If objRS.Fields("modalidade2").properties("IsAutoIncrement") = False Then
    objRS("modalidade2") = Trim(Request.Form("modalidade2"))
  End If
  If objRS.Fields("modalidade3").properties("IsAutoIncrement") = False Then
    objRS("modalidade3") = Trim(Request.Form("modalidade3"))
  End If
  If objRS.Fields("modalidade4").properties("IsAutoIncrement") = False Then
    objRS("modalidade4") = Trim(Request.Form("modalidade4"))
  End If
  If objRS.Fields("modalidade5").properties("IsAutoIncrement") = False Then
    objRS("modalidade5") = Trim(Request.Form("modalidade5"))
  End If
  If objRS.Fields("nome").properties("IsAutoIncrement") = False Then
    objRS("nome") = Trim(Request.Form("nome"))
  End If
  If objRS.Fields("nome_mae").properties("IsAutoIncrement") = False Then
    objRS("nome_mae") = Trim(Request.Form("nome_mae"))
  End If
  If objRS.Fields("nome_pai").properties("IsAutoIncrement") = False Then
    objRS("nome_pai") = Trim(Request.Form("nome_pai"))
  End If
  If objRS.Fields("numero").properties("IsAutoIncrement") = False Then
    objRS("numero") = Trim(Request.Form("numero"))
  End If
  If objRS.Fields("passaport").properties("IsAutoIncrement") = False Then
    objRS("passaport") = Trim(Request.Form("passaport"))
  End If
  If objRS.Fields("rg").properties("IsAutoIncrement") = False Then
    objRS("rg") = Trim(Request.Form("rg"))
  End If
  If objRS.Fields("telefone").properties("IsAutoIncrement") = False Then
    objRS("telefone") = Trim(Request.Form("telefone"))
  End If
   If objRS.Fields("parentesco").properties("IsAutoIncrement") = False Then
    objRS("parentesco") = Trim(Request.Form("parentesco"))
  End If
  If objRS.Fields("categoria1").properties("IsAutoIncrement") = False Then
    objRS("categoria1") = Trim(Request.Form("categoria1"))
  End If
  If objRS.Fields("categoria2").properties("IsAutoIncrement") = False Then
    objRS("categoria2") = Trim(Request.Form("categoria2"))
  End If
  If objRS.Fields("categoria3").properties("IsAutoIncrement") = False Then
    objRS("categoria3") = Trim(Request.Form("categoria3"))
  End If
  On Error Resume Next
  objRS.UpdateBatch
  objRS.Close
  Set objRS = Nothing
  objCon.Close
  Set objCon = Nothing
%>
  <%
Else
  If Not IsEmpty(Request.Form("recordno")) Then
    Set objCon = Server.CreateObject("ADODB.Connection")
    objCon.Open strCon

    Set objRS = Server.CreateObject("ADODB.Recordset")
    objRS.CursorLocation = 2
    objRS.CursorType = 0
    objRS.LockType = 3


  strQ = Request.Form("strQ")
  indice = Trim(Request.Form("indice"))
  If indice <> "" Then strQ = " SELECT * FROM inc_alunos WHERE " & indice

    objRS.Open strQ, objCon, , , &H0001
  If indice = "" Then objRS.Move Request.Form("recordno") - 1
%>
<span class="Estilo10"><BR>
  
  <INPUT type=hidden name=recordno value="<%=Request.Form("recordno")%>">
  <INPUT type=hidden name=strQ value="<%=Request.Form("strQ")%>">
  <INPUT type="hidden" name="indice" value="<%=indice%>">
  <%end if%>
  <%end if%>

<table width="880" border="1" align="center" cellspacing="0" bordercolor="#000000">
  <tr>
        <th width="236" valign="bottom" bgcolor="#FFFFFF" scope="row"><img src="imagens/logo1.png" width="77" height="97"><br>
        CEARÁ COLEGIAL</th>
        <th colspan="2" bgcolor="#FFFFFF" scope="row">FICHA DE INSCRIÇÃO ALUNO/ATLETA </th>
        <td width="147" rowspan="4"><div align="center"><img src="<%=(objRS.Fields.Item("foto").Value)%>" border=0 width=115 height="156"></div></td>
  </tr>
      
      <tr bordercolor="#000000">
        <th height="27" colspan="3" bgcolor="#000000" scope="row"><div align="center"><span class="Estilo19">IDENTIFICAÇÃO DA ESCOLA</span></div></th>
      </tr>
      <tr>
        <th height="38" colspan="3" bordercolor="#000000" scope="row"cellspacing="1"><div align="left">ESCOLA<span class="Estilo16">:<span class="Estilo18"><%=(objRS.Fields.Item("escola").Value)%></span></span></div></th>
      </tr>
      <tr>
        <th height="15%" colspan="2" bordercolor="#000000" scope="row"><div align="left">CNPJ<span class="Estilo16">:</span><span class="Estilo18"><%=(objRS.Fields.Item("cnpj_escola").Value)%></span></div></th>
        <th width="404" bordercolor="#000000"><div align="left">REG.ESCOLA:<span class="Estilo18"><%=(objRS.Fields.Item("id_escola").Value)%></span></div></th>
      </tr>
</table>
    <table width="880" border="1" align="center" cellspacing="0" bordercolor="#000000">
      <tr>
        <th colspan="3" bgcolor="#000000" scope="col"><span class="Estilo19">IDENTIFICAÇÃO DO ALUNO / ATLETA </span></th>
      </tr>
      <tr>
        <th colspan="2" scope="row"><div align="left">NOME COMPLETO: <span class="Estilo18"><%=(objRS.Fields.Item("nome").Value)%></span></div></th>
        <th width="207"><div align="left">APELIDO:<span class="Estilo18"><%=(objRS.Fields.Item("Id_alunos").Value)%></span></div></th>
      </tr>
      <tr>
        <th width="288" scope="row"><div align="left">NASCIMENTO:<span class="Estilo18"><%=(objRS.Fields.Item("data_nascimento").Value)%></span></div></th>
        <th width="369"><div align="left">SEXO:<span class="Estilo18"><%=(objRS.Fields.Item("sexo").Value)%></span></div></th>
        <th><div align="left">NATURALIDADE:<span class="Estilo18"><%=(objRS.Fields.Item("natural").Value)%></span></div></th>
      </tr>
      <tr>
        <th scope="row"><div align="left">IDENTIDADE:<span class="Estilo18"><%=(objRS.Fields.Item("rg").Value)%></span></div></th>
        <th><div align="left">MAT. ESCOLAR:<span class="Estilo18"><%=(objRS.Fields.Item("matricula_escola").Value)%></span></div></th>
        <th><div align="left">TELEFONE:<span class="Estilo18"><%=(objRS.Fields.Item("telefone").Value)%></span></div></th>
      </tr>
      <tr>
        <th scope="row"><div align="left">E-MAIL:<span class="Estilo27"><%=(objRS.Fields.Item("email").Value)%></span></div></th>
        <th scope="row"><div align="left">PASSAPORT:<span class="Estilo18"><%=(objRS.Fields.Item("passaport").Value)%></span></div></th>
        <th scope="row"><div align="left">CPF:<span class="Estilo18"><%=(objRS.Fields.Item("cpf").Value)%></span></div></th>
      </tr>
    </table>
    <table width="880" border="1" align="center" cellspacing="0" bordercolor="#000000">
      <tr>
        <th colspan="5" bgcolor="#000000" class="Estilo19" scope="col">DADOS RESPONSÁVEL LEGAL DO ALUNO/ATLETA</th>
      </tr>
      <tr>
        <th colspan="5" scope="row"><div align="left">NOME RESPONSÁVEL: <span class="Estilo18"><%=(objRS.Fields.Item("responsavel").Value)%></span></div></th>
      </tr>
      <tr>
        <th colspan="3" scope="row"><div align="left">RUA /AV: <span class="Estilo18"><%=(objRS.Fields.Item("endereco").Value)%></span></div></th>
        <th colspan="2"><div align="left">Nº<span class="Estilo18"><%=(objRS.Fields.Item("numero").Value)%></span></div></th>
      </tr>
      <tr>
        <th colspan="2" scope="row"><div align="left">COMPLEMENTO:</div></th>
        <th colspan="3"><div align="left">BAIRRO:<span class="Estilo18"><%=(objRS.Fields.Item("bairro").Value)%></span></div></th>
      </tr>
      <tr>
        <th width="311" scope="row"><div align="left">CEP:<span class="Estilo18"><%=(objRS.Fields.Item("cep").Value)%></span></div></th>
        <th colspan="2"><div align="left">MUNICIPIO:<span class="Estilo18"><%=(objRS.Fields.Item("cidade").Value)%></span></div></th>
        <th colspan="2"><div align="left">ESTADO: <span class="Estilo18">Ceará </span></div></th>
      </tr>
      <tr>
        <th scope="row"><div align="left">FONE:<span class="Estilo18"><%=(objRS.Fields.Item("telefone").Value)%></span></div></th>
        <th colspan="3"><div align="left">CELULAR:<span class="Estilo18"><%=(objRS.Fields.Item("telefone").Value)%></span></div></th>
        <th width="162"><div align="left">PARENTESCO:<span class="Estilo18"><%=(objRS.Fields.Item("parentesco").Value)%></span></div></th>
      </tr>
    </table>
    <table width="880" border="1" align="center" cellspacing="0" bordercolor="#000000">
      <tr>
        <th colspan="4" bgcolor="#000000" class="Estilo19" scope="col">OUTROS DADOS </th>
      </tr>
      <tr>
        <th height="28" colspan="4" scope="row"><div align="left"><span class="Estilo30">MODALIDADES:</span><span class="Estilo20"><span class="Estilo18"><%=(objRS.Fields.Item("modalidade1").Value)%>---<%=(objRS.Fields.Item("modalidade2").Value)%>---<%=(objRS.Fields.Item("modalidade3").Value)%>---<%=(objRS.Fields.Item("modalidade4").Value)%>---<%=(objRS.Fields.Item("modalidade5").Value)%></span></span></div></th>
      </tr>
      <tr>
        <th height="31" colspan="3" scope="row"><div align="left">CATEGORIA:<span class="Estilo18"><%=(objRS.Fields.Item("categoria1").Value)%></span>---<span class="Estilo18"><%=(objRS.Fields.Item("categoria2").Value)%></span>---<span class="Estilo18"><%=(objRS.Fields.Item("categoria3").Value)%></span></div></th>
        <td width="1" rowspan="4">&nbsp;</td>
      </tr>
      <tr>
        <th width="487" rowspan="3" valign="top" scope="row"> <span class="Estilo23">Carimbo e assinatura do secretário escolar</span><br></th>
        <td height="53" colspan="2" valign="top" class="Estilo23"> Coordenador de Esportes (nome legível):</td>
      </tr>
      <tr>
        <td height="78" colspan="2" valign="top"> <span class="Estilo23">Coordenador de Esportes (assinatura):</span><br></td>
      </tr>
      <tr>
        <td width="216" height="32"> Recebido em:<br>
          <span class="Estilo19">--------------</span>______/______/_____</td>
        <td width="158" valign="top"> Recebido por:<br></td>
      </tr>
    </table>
    <table width="880" height="169" border="1" align="center" cellspacing="0" bordercolor="#000000">
      <tr>
        <th height="30" bgcolor="#000000" scope="col"> <span class="Estilo19">TERMO DE ACEITAÇÃO </span></th>
      </tr>
      <tr>
        <th height="137" scope="row"><p align="justify"><span class="Estilo23"><span class="Estilo16">Os participantes concordam em autorizar, por 4 (quatro) anos, contados da apuração da equipe vencedora, o uso,
        pela Federação Cearense de Cultura e Desporto Escolar e pelo Grupo O POVO de Comunicação, de suas imagens,
        sons de voz e de seus nomes em filmes, vídeos, fotos, site, anúncios em jornais e revistas para divulgação do
        INTERCOLEGIAL O POVO, sem nenhum ônus para os mesmos.
        Para todos os efeitos legais, os participantes do presente evento declaram que as informações transmitidas no ato
        das inscrições são legítimas, responsabilizando-se e isentando as empresas promotoras de qualquer reclamação
        ou demanda que porventura venha a ser apresentada em juízo ou fora dele.</span></span></p>
          <p align="justify">______________________________________________________________________<br>
            Ass. Responsável
          Legal do Aluno<br>
          <br>
          <br>
          Data Geração Documento:<span class="Estilo18"> <%= now() %></span><br>
          </p></th>
      </tr>
    </table>   
<p align="center"><span class="Estilo16"><strong>FEDERAÇÃO CEARENSE DE CULTURA E DESPORTO ESCOLAR - CEARÁ COLEGIAL</strong></span><strong><br>
  CNPJ: 12.835.382/0001-52</strong></p>
<p align="center"><span class="Estilo5"><a href="javascript:DoPrinting()"><img src="imagens/icoImprimir.jpg" width="64" height="64" border="0" /></a>
  <script language="JavaScript1.2">
<!--
function DoPrinting(){
if (!window.print){
alert("Use o Netscape  ou Internet Explorer \n nas versões 4.0 ou superior!")
return
}
window.print()
}
//-->
</script>
</form>
</span><br>
<a href="http://www.sistemainc.com.br/cearacolegial/sistema/Cad_alunos/df_consulta.asp">VOLTAR </a><br>
</p>
<div align="center"><br>
</div>
</BODY>
</HTML>
