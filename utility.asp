<%@Language=VBSCript%>
<%Response.Buffer=True%>
<!DOCTYPE html>
<head>
	<meta charset="UTF-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
	<title>Sistema de Incrição  - Eventos</title>
	<meta name="description" content="">
	<meta name="author" content="">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<link rel="shortcut icon" href="/favicon.ico">
	<link rel="stylesheet" href="css/style.css?v=2">
	<body>
	<div id="container">
        <div id="contact-form" class="clearfix">
            <h1 align="center"> <img src="imagens/logo1.png" width="169" height="209"><br>
            CEARÁ COLEGIAL</h1>
            <h2 align="center">Sistema de Inscrição  - Eventos</h2>
	<%
method = Request.QueryString("method")

Select Case method
  Case "unauthorized"
    Unauthorized()
  Case "expired"
    Expired()
  Case "abandon"
    Abandon()
End Select


Sub Unauthorized()
  Response.Write "Voce nao tem autorizacao para  esta pagina!"
  Response.Write "<p><a href=""http://www.sistemainc.com.br/cearacolegial/sistema/"">Voltar</a>"
End Sub

Sub Expired()
Response.Write("<BR><B><p align=center ><strong>Atenção!</strong></B><BR><BR></p><p  align=center class=Estilo1 >A sua Escola nao entregou a ficha de Cadastrado! entre em contato com a Ceará Colegial <BR><A href= javascript:history.go(-1)>Clique aqui</a> encerrar</A><BR><BR></p>")
  'Response.Write "Sistema de registro de encerrado"
  'Response.Write "<p>Contate a Ceara Colegial "
  Session.Abandon
End Sub

Sub Abandon()
  Session.Abandon
  Response.Redirect "http://www.sistemainc.com.br/cearacolegial/sistema/"
End Sub
%>
      
      </div>
    </div>
	</body>
</html>