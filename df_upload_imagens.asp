<% @ LANGUAGE="VBSCRIPT" %>

<HTML>
<HEAD>
<TITLE>Imagens</TITLE>
<meta name="copyright" content="Dataform">
<meta name="keywords" content="dataform, asp dataform, aspdataform, asp-dataform">
<meta name="robots" content="ALL">
<style type="text/css">
<!--
.campo_alerta
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
border: 1px solid black;
background-color: #ffff99;
}
.texto_pagina
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
color: dimgray;
}

.tabela_formulario
{
width: 200;
background-color: white;
}

.titulo_campos
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
color: dimgray;
background-color: whitesmoke;
}

.campos_formulario
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
color: dimgray;
background-color: gainsboro;
border: 1px inset;
}

.botao_enviar
{
font-family: Tahoma, Verdana, Arial;
font-size: 11px;
color: white;
background-color: gray;
}
.Estilo1 {
	color: #FF0000;
	font-weight: bold;
	font-style: italic;
}
.Estilo4 {color: #000000; font-weight: bold; font-style: italic; }
-->
</style>
<script language=javascript>
function envia_imagem(imagem) {
self.opener.recebe_imagem('<%=Request("campo")%>', imagem);
window.close();
}
</script>
</HEAD>
<BODY class="texto_pagina">

<%
Sub BuildUploadRequest(RequestBin)
	PosBeg = 1
	PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
	boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
	boundaryPos = InstrB(1,RequestBin,boundary)
	Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
		Dim UploadControl
		Set UploadControl = CreateObject("Scripting.Dictionary")
		Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
		Pos = InstrB(Pos,RequestBin,getByteString("name="))
		PosBeg = Pos+6
		PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
		Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
		PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
		PosBound = InstrB(PosEnd,RequestBin,boundary)
		If  PosFile<>0 AND (PosFile<PosBound) Then
			PosBeg = PosFile + 10
			PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
			FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			UploadControl.Add "FileName", FileName
			Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
			PosBeg = Pos+14
			PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
			ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			UploadControl.Add "ContentType",ContentType
			PosBeg = PosEnd+4
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
		Else
			Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
			PosBeg = Pos+4
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
		End If
	UploadControl.Add "Value" , Value	
	UploadRequest.Add name, UploadControl	
	BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
	Loop

End Sub

Function getByteString(StringStr)
 For i = 1 to Len(StringStr)
 	char = Mid(StringStr,i,1)
	getByteString = getByteString & chrB(AscB(char))
 Next
End Function

Function getString(StringBin)
 getString =""
 For intCount = 1 to LenB(StringBin)
	getString = getString & chr(AscB(MidB(StringBin,intCount,1))) 
 Next
End Function

pasta_imagens = "../../Fotos/" & Request("pasta")
'pasta_imagens = "imagens/" & Request("pasta")
pasta = Server.URLEncode(Request("pasta"))
campo = Server.URLEncode(Request("campo"))

Set objFS = Server.CreateObject("Scripting.FileSystemObject")
If Not objFS.FolderExists(Server.MapPath(pasta_imagens)) Then
  objFS.CreateFolder(Server.MapPath(pasta_imagens))
End if

If Request("enviar") <> "" Then
  Set objFS = Nothing
  byteCount = Request.TotalBytes
  RequestBin = Request.BinaryRead(byteCount)
  Dim UploadRequest
  Set UploadRequest = CreateObject("Scripting.Dictionary")
  BuildUploadRequest  RequestBin
  contentType = UploadRequest.Item("blob").Item("ContentType")
  filepathname = UploadRequest.Item("blob").Item("FileName")
  filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
  value = UploadRequest.Item("blob").Item("Value")
  If Lcase(Right(filename,3)) = "jpg" Or Lcase(Right(filename,3)) = "gif" then
    Set objFS = Server.CreateObject("Scripting.FileSystemObject")
	If objFS.FileExists( Server.mappath(pasta_imagens & "\" & filename)) Then
%>

<script language=javascript>
alert("Erro ao enviar imagem, o arquivo '<%=filename%>' já existe, renomei e envie de novo")
enviar.disabled = false;
</script>

<%

	Else
	  If LenB(value) > 200000 then

%>

<script language=javascript>
alert("Erro ao enviar a imagem, o tamanho do arquivo deve ser menor que 200Kb")
enviar.disabled = false;
</script>

<%
      Else
%>
<strong>Aguarde o envio da imagem...</strong><br>

<input name="progress" value="0% enviado" style="border:none">
<table width="100" border="0" cellspacing="0" cellpadding="0" style="border: 1px inset">
  <tr>
    <td><input name="barra" style="border:none; background-color: orangered; height: 10; width:1" readonly=""></td>
    <td></td>
  </tr>
</table>

<%
      Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
      Set MyFile = ScriptObject.CreateTextFile( Server.mappath(pasta_imagens & "\" & filename))
	  progress = 0
	  n = 0
      For i = 1 to LenB(value)
        MyFile.Write chr(AscB(MidB(value,i,1)))
		progress = Fix((i * 100) / LenB(value))
		If n <> progress then
		  n = progress
%>

<script language=javascript>progress.value = "<%=n%>% enviado"</script>
<script language=javascript>barra.style.width = "<%=n%>"</script>

<%
        End if
      Next
      MyFile.Close
%>
<script language=javascript>
envia_imagem('<%=pasta_imagens & "/" & filename%>');
</script>

<%

	End If
	Set objFS = Nothing
  End if

Else
%>

<script language=javascript>
alert("Erro ao enviar a imagem, lembre-se que ela deve possuir extensão JPG ou GIF");
enviar.disabled = false;
</script>

<%
 End If
End If
%>

<FORM METHOD="post" ENCTYPE="multipart/form-data" ACTION="<%=Request.ServerVariables("SCRIPT_NAME")%>?campo=<%=campo%>&pasta=<%=pasta%>&enviar=sim" onSubmit="enviar.disabled=true">
Enviar uma nova imagem<BR>
<INPUT type="file" name="blob" class="campos_formulario" style="width: 100%"><BR>
<INPUT type="submit" name="enviar" value="Enviar" class="botao_enviar"><br>
<span class="Estilo1">(A imagem deve ter nó máximo 200Kb)</span>
</FORM>
<BR>
<DIV class="titulo_campos" style="width:100%; height:175px; visibility: visible; overflow: auto; border:1px solid">
  <p class="Estilo4">A imagem deve ter n&oacute; m&aacute;ximo 200Kb<br>
      <em><strong>Caso a Imagen ja exista renomear<br>
  a mesma para depois enviar</strong></em> </p>
</DIV>
</BODY>
</HTML>
