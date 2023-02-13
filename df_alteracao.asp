<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level1.inc"-->
<%
If session("allow") = False Then Response.Redirect "../index.asp"
If session("clearance") < 1 Then Response.Redirect "utility.asp?method=unauthorized"

pagina_consulta = "df_consulta.asp"
pagina_alteracao = "df_alteracao.asp"
pagina_inclusao = "df_inclusao.asp"
pagina_login = "df_login.asp"

%>

<HTML>
<HEAD>
<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<meta name="robots" content="ALL">
<title>Alterar Alunos</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
	<link rel="stylesheet" type='text/css' media='all' href="static/css/base.css">
    <link rel="stylesheet" type='text/css' media='all' href="static/css/colors.css">
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.3/jquery.min.js" type="text/javascript"></script>

	
<script src="maskedinput.js" type="text/javascript"></script>
<script type="text/javascript"> 
jQuery.noConflict();
jQuery(function($){   //$("#data").mask("99/99/9999");   $("#data").mask("*9/99/9999",{completed:function(){alert("You typed the following: "+this.val());}});
 $("#data_nascimento").mask("99/99/9999",{placeholder:" "});   $("#telefone").mask("(099) 99999-9999");   $("#cpf").mask("999.999.999-99");   $("#cep1").mask("99999 - 999");    $("#cnpj").mask("99.999.999/9999-99");   $("#placa").mask("aaa - 9999"); });</script>
 
<script type="text/javascript">
src="https://ajax.googleapis.com/ajax/libs/jquery/1.5.0/jquery.min.js"
function getEndereco() {
	// Se o campo CEP não estiver vazio
	if($.trim($("#cep").val()) != ""){
	$.getScript("http://cep.republicavirtual.com.br/web_cep.php?formato=javascript&cep="+$("#cep").val(), function(){
	if (resultadoCEP["tipo_logradouro"] != '') {
		if (resultadoCEP["resultado"]) {
		// troca o valor dos elementos
			$("#endereco").val(unescape(resultadoCEP["tipo_logradouro"]) + " " + unescape(resultadoCEP["logradouro"]));
			$("#bairro").val(unescape(resultadoCEP["bairro"]));
			$("#cidade").val(unescape(resultadoCEP["cidade"]));
			$("#estado").val(unescape(resultadoCEP["uf"]));
			//$("#numero").focus();
			}
		}		
	});
	}
}
</script>
<script language="JavaScript" type="text/javascript">
    function maiuscula(id) {
                //palavras para ser ignoradas
                var wordsToIgnore = ["DOS", "DAS", "de", "do", "Dos", "Das"],
                minLength = 3;
                var str = $(id).val(); 
                var getWords = function(str) {
                    return str.match(/\S+\s*/g);
                }
                $(id).each(function () {
                    var words = getWords(this.value);
                    $.each(words, function (i, word) {
                        // somente continua se a palavra nao estiver na lista de ignorados
                        if (wordsToIgnore.indexOf($.trim(word)) == -1 && $.trim(word).length > minLength) {
                            words[i] = words[i].charAt(0).toUpperCase() + words[i].slice(1).toLowerCase();
                        } else{
                        words[i] = words[i].toLowerCase();}
                    });
                    this.value = words.join("");
                });
            };
	</script>
<script language="JavaScript" type="text/javascript">
<!--
function abre_janela(width, height, nome) {
var top; var left;
top = ( (screen.height/2) - (height/2) )
left = ( (screen.width/2) - (width/2) )
window.open('',nome,'width='+width+',height='+height+',scrollbars=no,toolbar=no,location=no,status=no,menubar=no,resizable=no,left='+left+',top='+top);
}
function recebe_imagem(campo, imagem){
var foto = 'img_' + campo
document.form_incluir[campo].value = imagem;
document.form_incluir[foto].src = imagem;
}
function verifica_form(form) {
var passed = false;
var ok = false
var campo
for (i = 0; i < form.length; i++) {
  campo = form[i].name;
  if (form[i].df_verificar == "sim") {
    if (form[i].type == "text"  | form[i].type == "textarea" | form[i].type == "select-one") {
      if (form[i].value == "" | form[i].value == "http://") {
		form[campo].className='campo_alerta'
        form[campo].focus();
        alert("Preencha corretamente o campo");
        return passed;
        stop;
      }
    }
    else if (form[i].type == "radio") {
      for (x = 0; x < form[campo].length; x++) {
        ok = false;
        if (form[campo][x].checked) {
          ok = true;
          break;
        }
      }
      if (ok == false) {
        form[campo][0].focus();
		form[campo][0].select();
        alert("Informe uma das opcões");
        return passed;
        stop;
      }
    }
    var msg = ""
    if (form[campo].df_validar == "cpf") msg = checa_cpf(form[campo].value);
    if (form[campo].df_validar == "cnpj") msg = checa_cnpj(form[campo].value);
    if (form[campo].df_validar == "cpf_cnpj") {
	  msg = checa_cpf(form[campo].value);
	  if (msg != "") msg = checa_cnpj(form[campo].value);
	}
    if (form[campo].df_validar == "email") msg = checa_email(form[campo].value);
    if (form[campo].df_validar == "numerico") msg = checa_numerico(form[campo].value);
    if (msg != "") {
	  if (form[campo].df_validar == "cpf_cnpj") msg = "informe corretamente o número do CPF ou CNPJ";
	  form[campo].className='campo_alerta'
      form[campo].focus();
      form[campo].select();
      alert(msg);
      return passed;
      stop;
    }
  }
}
passed = true;
return passed;
}
function desabilita_cor(campo) {
campo.className='campos_formulario'
}
function checa_numerico(String) {
var mensagem = "Este campo aceita somente números"
var msg = "";
if (isNaN(String)) msg = mensagem;
return msg;
}
function checa_email(campo) {
var mensagem = "Informe corretamente o email"
var msg = "";
var email = campo.match(/(\w+)@(.+)\.(\w+)$/);
if (email == null){
  msg = mensagem;
  }
return msg;
}
function checa_cpf(CPF) {
var mensagem = "informe corretamente o número do CPF"
var msg = "";
if (CPF.length != 11 || CPF == "00000000000" || CPF == "11111111111" ||
  CPF == "22222222222" ||	CPF == "33333333333" || CPF == "44444444444" ||
  CPF == "55555555555" || CPF == "66666666666" || CPF == "77777777777" ||
  CPF == "88888888888" || CPF == "99999999999")
msg = mensagem;
soma = 0;
for (y=0; y < 9; y ++)
soma += parseInt(CPF.charAt(y)) * (10 - y);
resto = 11 - (soma % 11);
if (resto == 10 || resto == 11)resto = 0;
if (resto != parseInt(CPF.charAt(9)))
  msg = mensagem; soma = 0;
for (y = 0; y < 10; y ++)
  soma += parseInt(CPF.charAt(y)) * (11 - y);
resto = 11 - (soma % 11);
if (resto == 10 || resto == 11) resto = 0;
if (resto != parseInt(CPF.charAt(10)))
  msg = mensagem;
return msg;
}
function checa_cnpj(s) {
var mensagem = "informe corretamente o número do CNPJ"
var msg = "";
var y;
var c = s.substr(0,12);
var dv = s.substr(12,2);
var d1 = 0;
for (y = 0; y < 12; y++)
{
d1 += c.charAt(11-y)*(2+(y % 8));
}
if (d1 == 0) msg = mensagem;
d1 = 11 - (d1 % 11);
if (d1 > 9) d1 = 0;
if (dv.charAt(0) != d1)msg = mensagem;
d1 *= 2;
for (y = 0; y < 12; y++)
{
d1 += c.charAt(11-y)*(2+((y+1) % 8));
}
d1 = 11 - (d1 % 11);
if (d1 > 9) d1 = 0;
if (dv.charAt(1) != d1) msg = mensagem;
return msg;
}
function mascara_data(data){ 
var mydata = ''; 
mydata = mydata + data; 
if (mydata.length == 2){ 
mydata = mydata + '/'; 
} 
if (mydata.length == 5){ 
mydata = mydata + '/'; 
} 
return mydata; 
} 
function verifica_data(data) { 
if (data.value != "") {
dia = (data.value.substring(0,2));
mes = (data.value.substring(3,5)); 
ano = (data.value.substring(6,10)); 
situacao = ""; 
if ((dia < 01)||(dia < 01 || dia > 30) && (  mes == 04 || mes == 06 || mes == 09 || mes == 11 ) || dia > 31) { 
situacao = "falsa"; 
} 
if (mes < 01 || mes > 12 ) { 
situacao = "falsa"; 
}
if (mes == 2 && ( dia < 01 || dia > 29 || ( dia > 28 && (parseInt(ano / 4) != ano / 4)))) { 
situacao = "falsa"; 
} 
if (situacao == "falsa") { 
data.focus();
data.select();
alert("Data inválida!"); 
}
} 
}
//-->
</script>
<style type="text/css">
<!--
.Estilo6 {font-family: tahoma}
.Estilo7 {	font-family: Tahoma;
	font-weight: bold;
}
.Estilo9 {font-size: 16px; font-family: Tahoma; font-weight: bold; }
.Estilo10 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.texto_pagina #Form1 table tr td p {
	text-align: center;
}
.Estilo3 {font-family: Arial, Helvetica, sans-serif}
assadasd {
	text-align: left;
}
-->
</style>
</HEAD>
<body>
	<section class="bg-primary">
          <span class="background dark" style="background-image:url('https://source.unsplash.com/RkBTPqPEGDo/')"></span> 	      
          <div class="wrap">
           <fieldset>
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
  If objRS.Fields("sexo").properties("IsAutoIncrement") = False Then
    objRS("sexo") = Trim(Request.Form("sexo"))
  End If
  If objRS.Fields("responsavel").properties("IsAutoIncrement") = False Then
    objRS("responsavel") = Trim(Request.Form("responsavel"))
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
   If objRS.Fields("matricula_escola").properties("IsAutoIncrement") = False Then
    objRS("matricula_escola") = Trim(Request.Form("matricula_escola"))
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

<BR>
<span class="Estilo10"><img src="imagens/logo1.png" width="122" height="149" align="center"><br>
</span>
<h2 align="center" class="Estilo3"><strong>Ceará Colegial informa</strong></h2>
<align="center" class="Estilo3">Atleta alterado<BR>O Atleta foi alterado com sucesso.<a href="df_consulta.asp">Voltar</a></span></strong><BR>
<BR>

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
</span><img src="imagens/logo1.png"class="aligncenter" width="122" height="149" >
<h2 class="aligncenter">ALTERAR DADOS ATLETAS</h2>
<p class="aligncenter">Sistema de inscrição de eventos<br>
  <a href="http://www.sistemainc.com.br/cearacolegial/sistema/Cad_alunos/df_consulta.asp">Voltar </a></p>
<BR>
<form name="form_incluir" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>" accept-charset="ISO-8859-1" id="Form1" onSubmit="return Validateform_incluir(this);">
<INPUT type=hidden name=recordno value="<%=Request.Form("recordno")%>">
<INPUT type=hidden name=strQ value="<%=Request.Form("strQ")%>">
<INPUT type="hidden" name="indice" value="<%=indice%>">

<table width="863" border="0" align="center">
  <tr>
    <td width="653" valign="top">
        <div align="center">
          <table width="653" border="0" cellpadding="0" cellspacing="0" bordercolor="#000000" bgcolor="#FFFFFF">
            <!--DWLayoutTable-->
            <tr align="left" valign="middle" bordercolor="#FFFFFF">
              <td width="167" height="31" valign="bottom" bgcolor="#FFFFFF" class="style11"><%If objRS.Fields("cnpj_escola").properties("IsAutoIncrement") = False Then%>
                <INPUT style="width=350" type="hidden" name="cnpj_escola" maxlength="255" value="<%=(objRS.Fields.Item("cnpj_escola").Value)%>" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
              <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("cnpj_escola").Value) & "</B>"
End If
%></td>
              <td width="501" valign="middle" bordercolor="#666666" bgcolor="#FFFFFF"><div align="center" class="Estilo7">
                <div align="left"></div>
              </div></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><span class="style13 style6 style8"><strong><font face="Tahoma">Nome</font></strong></span></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><%If objRS.Fields("nome").properties("IsAutoIncrement") = False Then%>
			  
 <INPUT name="nome" type="text"  value="<%=(objRS.Fields.Item("nome").Value)%>" class="t1 required-entry firstname" id="billing:firstname"  size="50" maxlength="255" onKeyUp="maiuscula('.firstname')">
                  <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("nome").Value) & "</B>"
End If
%>
                  <%If objRS.Fields("escola").properties("IsAutoIncrement") = False Then%>
                  <INPUT style="width=350" type="hidden" name="escola" maxlength="255" value="<%=(objRS.Fields.Item("escola").Value)%>" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
              <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("escola").Value) & "</B>"
End If
%></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><strong><font face="Tahoma">Sexo</font></strong></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><label>
                <%If objRS.Fields("sexo").properties("IsAutoIncrement") = False Then%>
                <select name="sexo" id="sexo">
                  <option value="Masculino">Masculino</option>
                  <option value="Feminino">Feminino</option>
                </select>
                <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("nome").Value) & "</B>"
End If
%>
              </label></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><strong><font face="Tahoma">Data Nascimento </font></strong></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><h1 align="left" class="style14 style5">
                <%If objRS.Fields("data_nascimento").properties("IsAutoIncrement") = False Then%>
                <INPUT style="width=350" type="text" name="data_nascimento" maxlength="255" value="<%=(objRS.Fields.Item("data_nascimento").Value)%>" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
                <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("data_nascimento").Value) & "</B>"
End If
%>
                <%If objRS.Fields("Id_alunos").properties("IsAutoIncrement") = False Then%>
                <INPUT style="width=350" type="hidden" name="Id_alunos" maxlength="255" value="<%=(objRS.Fields.Item("Id_alunos").Value)%>" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
 <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("Id_alunos").Value) & "</B>"
End If
%>
</h1></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><span class="style13 style6 style8"><strong><font face="Tahoma">E-mail</font></strong></span></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><h1 align="left" class="style14 style5">
                <%If objRS.Fields("email").properties("IsAutoIncrement") = False Then%>
                <INPUT name="email" type="text" class=campos_formulario style="width=350" onKeyPress="desabilita_cor(this)" value="<%=(objRS.Fields.Item("email").Value)%>" size="50" maxlength="255"  df_verificar="sim">
                <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("email").Value) & "</B>"
End If
%>
</h1></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><strong><font face="Tahoma">Endere&ccedil;o</font></strong></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><span class="style14">
                <%If objRS.Fields("endereco").properties("IsAutoIncrement") = False Then%>
                <INPUT name="endereco" type="text" class=campos_formulario style="width=350" onKeyPress="desabilita_cor(this)" value="<%=(objRS.Fields.Item("endereco").Value)%>" size="55" maxlength="255"  df_verificar="sim">
                <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("endereco").Value) & "</B>"
End If
%>
              </span></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><strong><font face="Tahoma">N&ordm;</font></strong></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><h1 class="style5">
                <%If objRS.Fields("numero").properties("IsAutoIncrement") = False Then%>
                  <INPUT name="numero" type="text" class=campos_formulario style="width=350" onKeyPress="desabilita_cor(this)" value="<%=(objRS.Fields.Item("numero").Value)%>" size="6" maxlength="50"  df_verificar="sim">
                  <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("numero").Value) & "</B>"
End If
%>
</h1></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><span class="style13 style6 style8"><strong><font face="Tahoma">Bairro</font></strong></span></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><h1 class="style5">
                <%If objRS.Fields("bairro").properties("IsAutoIncrement") = False Then%>
                <INPUT style="width=350" type="text" name="bairro" maxlength="255" value="<%=(objRS.Fields.Item("bairro").Value)%>" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
                <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("bairro").Value) & "</B>"
End If
%>
</h1></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><strong><font face="Tahoma">Cidade</font></strong></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><span class="style14 style5">
                <%If objRS.Fields("cidade").properties("IsAutoIncrement") = False Then%>
                <INPUT style="width=350" type="text" name="cidade" maxlength="255" value="<%=(objRS.Fields.Item("cidade").Value)%>" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
                <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("cidade").Value) & "</B>"
End If
%>
              </span></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><span class="Estilo9">Telefone</span></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><h1 align="left" class="style14 style5">
                <%If objRS.Fields("telefone").properties("IsAutoIncrement") = False Then%>
<INPUT style="width=350" type="text" name="telefone" maxlength="255" value="<%=(objRS.Fields.Item("telefone").Value)%>" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
<%
Else
  Response.Write "<B>" & (objRS.Fields.Item("telefone").Value) & "</B>"
End If
%>
              </h1></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><span class="style13 style6 style8"><strong><font face="Tahoma">Cep</font></strong></span></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><h1 align="left" class="style14 style5">
                <%If objRS.Fields("cep").properties("IsAutoIncrement") = False Then%>
                <INPUT style="width=350" type="text" name="cep" maxlength="255" value="<%=(objRS.Fields.Item("cep").Value)%>" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
                <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("cep").Value) & "</B>"
End If
%>
</h1></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><span class="Estilo9">RG</span></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><span class="style14 style5">
              <%If objRS.Fields("rg").properties("IsAutoIncrement") = False Then%>
              <INPUT style="width=350" type="text" name="rg" maxlength="255" value="<%=(objRS.Fields.Item("rg").Value)%>" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
              <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("rg").Value) & "</B>"
End If
%>
              </span></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><span class="style13 style6 style8"><strong><font face="Tahoma">CPF</font></strong></span></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><h1 align="left" class="style14 style5">
                <%If objRS.Fields("cpf").properties("IsAutoIncrement") = False Then%>
                <INPUT style="width=350" type="text" name="cpf" maxlength="255" value="<%=(objRS.Fields.Item("cpf").Value)%>" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
                <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("cpf").Value) & "</B>"
End If
%>
              </h1></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><strong><font face="Tahoma">Responsável</font></strong></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><span class="style14 style5">
                <%If objRS.Fields("responsavel").properties("IsAutoIncrement") = False Then%>
                <INPUT name="responsavel" type="text" class=campos_formulario id="responsavel" style="width=350" onKeyPress="desabilita_cor(this)" value="<%=(objRS.Fields.Item("responsavel").Value)%>" size="50" maxlength="255"  df_verificar="sim">
                <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("responsavel").Value) & "</B>"
End If
%>
              </span></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><span class="style13 style6 style8"><strong><font face="Tahoma">Parentesco</font></strong></span></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><h1 align="left" class="style14 style5">
                <%If objRS.Fields("parentesco").properties("IsAutoIncrement") = False Then%>
                <INPUT name="parentesco" type="text" class=campos_formulario id="parentesco" style="width=350" onKeyPress="desabilita_cor(this)" value="<%=(objRS.Fields.Item("parentesco").Value)%>" size="50" maxlength="255"  df_verificar="sim">
                <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("parentesco").Value) & "</B>"
End If
%>
</h1></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF"><span class="style13 style6 style8"><strong><font face="Tahoma">Matricula</font></strong></span></td>
              <td align="left" valign="bottom" bgcolor="#FFFFFF" class="even"><h1 align="left" class="style14 style5">
                  <label></label>
                  <%If objRS.Fields("matricula_escola").properties("IsAutoIncrement") = False Then%>
                  <INPUT name="matricula_escola" type="text" class=campos_formulario id="matricula_escola" style="width=350" onKeyPress="desabilita_cor(this)" value="<%=(objRS.Fields.Item("matricula_escola").Value)%>" size="50" maxlength="255"  df_verificar="sim">
                  <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("nome_mae").Value) & "</B>"
End If
%>
</h1></td>
            </tr>
            <tr>
              <td align="left" valign="middle" bgcolor="#FFFFFF" class="Estilo6 style17"><strong>Passaport</strong></td>
              <td align="left" valign="middle" bgcolor="#FFFFFF" class="bg-lightgray"><%If objRS.Fields("passaport").properties("IsAutoIncrement") = False Then%>
                <INPUT style="width=350" type="text" name="passaport" maxlength="255" value="<%=(objRS.Fields.Item("passaport").Value)%>" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
              <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("passaport").Value) & "</B>"
End If
%></td>
            </tr>
            <tr align="left" valign="bottom">
              <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21 Estilo7">Modalidade 01 </td>
              <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21"><label>
              <%If objRS.Fields("modalidade1").properties("IsAutoIncrement") = False Then%>
              <select name="modalidade1">
                <option value="<%=(objRS.Fields.Item("modalidade1").Value)%>"><%=(objRS.Fields.Item("modalidade1").Value)%></option>
                <option value="Nao Tem">Nâo Tem</option>
       	        <option value="ATLETISMO">ATLETISMO</option>
		        <option value="JUD&Ocirc;">JUD&Ocirc;</option>
		        <option value="NATA&Ccedil;&Atilde;O">NATA&Ccedil;&Atilde;O</option>
		        <option value="XADREZ">XADREZ</option>
		        <option value="VOLEIBOL">VOLEIBOL</option>
		        <option value="HANDEBOL">HANDEBOL</option>
		        <option value="FUTSAL">FUTSAL</option>
                <option value="FUTEBOL DE SALAO">FUTEBOL DE SALAO</option>
		        <option value="BASQUETE">BASQUETE</option>
		        <option value="DAMA">DAMA</option>
		        <option value="HAND BEACH">HAND BEACH</option>
		        <option value="VOLEI DE PRAIA">VOLEI DE PRAIA</option>
		        <option value="TAEKWONDO">TAEKWONDO</option>
		        <option value="TENIS DE MESA">TENIS DE MESA</option>
		        <option value="CICLISMO">CICLISMO</option>
		        <option value="BADMINTON">BADMINTON</option>
		        <option value="KARATE">KARATE</option>
              </select>
              <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("modalidade1").Value) & "</B>"
End If
%>
</label></td>
            </tr>
            <tr align="left" valign="bottom">
              <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21 Estilo7">Modalidade 02 </td>
              <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21"><%If objRS.Fields("modalidade2").properties("IsAutoIncrement") = False Then%>
                <select name="modalidade2" id="modalidade2">
                  <option value="<%=(objRS.Fields.Item("modalidade2").Value)%>"><%=(objRS.Fields.Item("modalidade2").Value)%></option>
				  <option value="Nao Tem">Nâo Tem</option>
       	        <option value="ATLETISMO">ATLETISMO</option>
		        <option value="JUD&Ocirc;">JUD&Ocirc;</option>
		        <option value="NATA&Ccedil;&Atilde;O">NATA&Ccedil;&Atilde;O</option>
		        <option value="XADREZ">XADREZ</option>
		        <option value="VOLEIBOL">VOLEIBOL</option>
		        <option value="HANDEBOL">HANDEBOL</option>
		        <option value="FUTSAL">FUTSAL</option>
                <option value="FUTEBOL DE SALAO">FUTEBOL DE SALAO</option>
		        <option value="BASQUETE">BASQUETE</option>
		        <option value="DAMA">DAMA</option>
		        <option value="HAND BEACH">HAND BEACH</option>
		        <option value="VOLEI DE PRAIA">VOLEI DE PRAIA</option>
		        <option value="TAEKWONDO">TAEKWONDO</option>
		        <option value="TENIS DE MESA">TENIS DE MESA</option>
		        <option value="CICLISMO">CICLISMO</option>
		        <option value="BADMINTON">BADMINTON</option>
		        <option value="KARATE">KARATE</option>
                </select>
              <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("modalidade2").Value) & "</B>"
End If
%></td>
            </tr>
            <tr align="left" valign="bottom">
              <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21 Estilo7">Modalidade 03</td>
              <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21"><%If objRS.Fields("modalidade3").properties("IsAutoIncrement") = False Then%>
                <select name="modalidade3" id="modalidade3">
                  <option value="<%=(objRS.Fields.Item("modalidade3").Value)%>"><%=(objRS.Fields.Item("modalidade3").Value)%></option>
				   <option value="Nao Tem">Nâo Tem</option>
       	        <option value="ATLETISMO">ATLETISMO</option>
		        <option value="JUD&Ocirc;">JUD&Ocirc;</option>
		        <option value="NATA&Ccedil;&Atilde;O">NATA&Ccedil;&Atilde;O</option>
		        <option value="XADREZ">XADREZ</option>
		        <option value="VOLEIBOL">VOLEIBOL</option>
		        <option value="HANDEBOL">HANDEBOL</option>
		        <option value="FUTSAL">FUTSAL</option>
                <option value="FUTEBOL DE SALAO">FUTEBOL DE SALAO</option>
		        <option value="BASQUETE">BASQUETE</option>
		        <option value="DAMA">DAMA</option>
		        <option value="HAND BEACH">HAND BEACH</option>
		        <option value="VOLEI DE PRAIA">VOLEI DE PRAIA</option>
		        <option value="TAEKWONDO">TAEKWONDO</option>
		        <option value="TENIS DE MESA">TENIS DE MESA</option>
		        <option value="CICLISMO">CICLISMO</option>
		        <option value="BADMINTON">BADMINTON</option>
		        <option value="KARATE">KARATE</option>
                </select>
              <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("modalidade3").Value) & "</B>"
End If
%></td>
            </tr>
            <tr align="left" valign="bottom">
              <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21 Estilo7">Modalidade 04</td>
              <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21"><%If objRS.Fields("modalidade4").properties("IsAutoIncrement") = False Then%>
                <select name="modalidade4" id="modalidade4">
                  <option value="<%=(objRS.Fields.Item("modalidade4").Value)%>"><%=(objRS.Fields.Item("modalidade4").Value)%></option>
				   <option value="Nao Tem">Nâo Tem</option>
       	        <option value="ATLETISMO">ATLETISMO</option>
		        <option value="JUD&Ocirc;">JUD&Ocirc;</option>
		        <option value="NATA&Ccedil;&Atilde;O">NATA&Ccedil;&Atilde;O</option>
		        <option value="XADREZ">XADREZ</option>
		        <option value="VOLEIBOL">VOLEIBOL</option>
		        <option value="HANDEBOL">HANDEBOL</option>
		        <option value="FUTSAL">FUTSAL</option>
                <option value="FUTEBOL DE SALAO">FUTEBOL DE SALAO</option>
		        <option value="BASQUETE">BASQUETE</option>
		        <option value="DAMA">DAMA</option>
		        <option value="HAND BEACH">HAND BEACH</option>
		        <option value="VOLEI DE PRAIA">VOLEI DE PRAIA</option>
		        <option value="TAEKWONDO">TAEKWONDO</option>
		        <option value="TENIS DE MESA">TENIS DE MESA</option>
		        <option value="CICLISMO">CICLISMO</option>
		        <option value="BADMINTON">BADMINTON</option>
		        <option value="KARATE">KARATE</option>
                </select>
              <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("modalidade4").Value) & "</B>"
End If
%></td>
            </tr>
            <tr align="left" valign="bottom">
              <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21 Estilo7">Modalidade 05</td>
              <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21"><%If objRS.Fields("modalidade5").properties("IsAutoIncrement") = False Then%>
                <select name="modalidade5" id="modalidade5">
                  <option value="<%=(objRS.Fields.Item("modalidade5").Value)%>"><%=(objRS.Fields.Item("modalidade5").Value)%></option>
                <option value="Nao Tem">Nâo Tem</option>
       	        <option value="ATLETISMO">ATLETISMO</option>
		        <option value="JUD&Ocirc;">JUD&Ocirc;</option>
		        <option value="NATA&Ccedil;&Atilde;O">NATA&Ccedil;&Atilde;O</option>
		        <option value="XADREZ">XADREZ</option>
		        <option value="VOLEIBOL">VOLEIBOL</option>
		        <option value="HANDEBOL">HANDEBOL</option>
		        <option value="FUTSAL">FUTSAL</option>
                <option value="FUTEBOL DE SALAO">FUTEBOL DE SALAO</option>
		        <option value="BASQUETE">BASQUETE</option>
		        <option value="DAMA">DAMA</option>
		        <option value="HAND BEACH">HAND BEACH</option>
		        <option value="VOLEI DE PRAIA">VOLEI DE PRAIA</option>
		        <option value="TAEKWONDO">TAEKWONDO</option>
		        <option value="TENIS DE MESA">TENIS DE MESA</option>
		        <option value="CICLISMO">CICLISMO</option>
		        <option value="BADMINTON">BADMINTON</option>
		        <option value="KARATE">KARATE</option>
                </select>
                <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("modalidade5").Value) & "</B>"
End If
%>
                <%If objRS.Fields("id_escola").properties("IsAutoIncrement") = False Then%>
                <INPUT style="width=350" type="hidden" name="id_escola" maxlength="255" value="<%=(objRS.Fields.Item("id_escola").Value)%>" onKeyPress="desabilita_cor(this)"  df_verificar="sim" class=campos_formulario>
              <%
Else
  Response.Write "<B>" & (objRS.Fields.Item("id_escola").Value) & "</B>"
End If
%></td>
            </tr>
            <tr align="left" valign="bottom">
            <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21 Estilo7">Categoria 01</td>
              <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21"><%If objRS.Fields("categoria1").properties("IsAutoIncrement") = False Then%>
                <select name="categoria1" id="categoria1">
                  <option value="<%=(objRS.Fields.Item("categoria1").Value)%>"><%=(objRS.Fields.Item("categoria1").Value)%></option>
                  <option value="Sub09">Sub09</option>
                  <option value="Sub11">Sub11</option>
                  <option value="Sub13">Sub13</option>
                  <option value="Sub17">Sub15</option>
                  <option value="Sub17">Sub17</option>
                  <option value="Sub18">Sub18</option>
                  
                </select>
                <%
Else
  Response.Write "<B>(Autom&aacute;tico)</B>"
End If
%>
              </strong></td>
            </tr>
            <tr align="left" valign="bottom">
            <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21 Estilo7">Categoria 02</td>
              <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21"><%If objRS.Fields("categoria2").properties("IsAutoIncrement") = False Then%>
                <select name="categoria2" id="categoria2">
                  <option value="<%=(objRS.Fields.Item("categoria2").Value)%>"><%=(objRS.Fields.Item("categoria2").Value)%></option>
                <option value="Nao Tem">Nâo Tem</option>
                  <option value="Sub09">Sub09</option>
                  <option value="Sub11">Sub11</option>
                  <option value="Sub13">Sub13</option>
                  <option value="Sub17">Sub15</option>
                  <option value="Sub17">Sub17</option>
                  <option value="Sub18">Sub18</option>
                 
                </select>
                <%
Else
  Response.Write "<B>(Autom&aacute;tico)</B>"
End If
%>
              </strong></td>
            </tr>
            <tr align="left" valign="bottom">
              <<td valign="baseline" bgcolor="#FFFFFF" class="Estilo21 Estilo7">Categoria 03</td>
              <td valign="baseline" bgcolor="#FFFFFF" class="Estilo21"><%If objRS.Fields("categoria3").properties("IsAutoIncrement") = False Then%>
                <select name="categoria3" id="categoria3">
                  <option value="<%=(objRS.Fields.Item("categoria3").Value)%>"><%=(objRS.Fields.Item("categoria3").Value)%></option>
                <option value="Nao Tem">Nâo Tem</option>
                  <option value="Sub09">Sub09</option>
                  <option value="Sub11">Sub11</option>
                  <option value="Sub13">Sub13</option>
                  <option value="Sub17">Sub15</option>
                  <option value="Sub17">Sub17</option>
                  
                </select>
                <%
Else
  Response.Write "<B>(Autom&aacute;tico)</B>"
End If
%>
              </strong></td>
              <%end if%>
            </tr>
          </table>
        </div>
      <p align="center">
        <input type="submit" name="salvar" value="Enviar" class=botao_enviar>
      </p></td>
    <td width="160" valign="top"><p>Click em Redefinir imagem caso nao apareça a foto.<br>
      <INPUT type="hidden" name="foto" value="<%=(objRS.Fields.Item("foto").Value)%>" df_verificar="sim">
  <a href="df_upload_imagens.asp?campo=<%=Server.URLEncode("foto")%>&pasta=<%=Server.URLEncode("imagens_foto")%>" onClick="abre_janela(250, 300, 'alterar_imagem')" target="alterar_imagem" class="texto_pagina">Alterar Imagem</a> &nbsp;|&nbsp;<a href="javascript: recebe_imagem('foto','<%=(objRS.Fields.Item("foto").Value)%>')" class="texto_pagina">Redefinir Imagem</a></p>
      <DIV id="Layer1" style="width: 1px; height:1px; visibility: visible; border:1px solid black"><img src="imagens/anonimo12.jpg" alt="" name="img_foto" width="381" height="472" border="0"></DIV></td>
  </tr>
</table>
<p>&nbsp; </p>
</form>

<%
    
  End If

%>
</fieldset>
</section>
</BODY>
</HTML>
