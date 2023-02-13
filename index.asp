<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Buffer = True%>
<!doctype html>
<html lang="en" prefix="og: http://ogp.me/ns#">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- SEO -->
    <title>Sistema de Inscrição - Eventos</title>
    <meta name="description" content="Sistema de Inscrição - Eventos - Ceara Colegial">    
    <link href="https://fonts.googleapis.com/css?family=Roboto:100,100i,300,300i,400,400i,700,700i%7CMaitree:200,300,400,600,700&amp;subset=latin-ext" rel="stylesheet">
    <link rel="stylesheet" type='text/css' media='all' href="../static/css/base.css">
    <link rel="stylesheet" type='text/css' media='all' href="../static/css/colors.css">
    <link rel="stylesheet" type='text/css' media='all' href="../static/css/svg-icons.css">
    <link rel="shortcut icon" sizes="16x16" href="../static/images/favicons/favicon.png">
     <!-- Android -->
    <meta name="mobile-web-app-capable" content="yes">
    <meta name="theme-color" content="#333333">
  <script src="http://ajax.googleapis.com/ajax/libs/jquery/1.3/jquery.min.js" type="text/javascript"></script>
  
  <!-- Mascara CNPJ -->
<script src="maskedinput.js" type="text/javascript"></script>
<script type="text/javascript"> 
jQuery.noConflict();
jQuery(function($){   //$("#data").mask("99/99/9999");   $("#data").mask("*9/99/9999",{completed:function(){alert("You typed the following: "+this.val());}});
 $("#data").mask("99/99/9999",{placeholder:" "});   $("#telefone").mask("(099) 9999-9999");   $("#cpf").mask("999.999.999-99");   $("#cep").mask("99999 - 999");    $("#cnpj").mask("99.999.999/9999-99");   $("#placa").mask("aaa - 9999"); });</script> 
  </head>
  
  <body>
   <section class="bg-primary">
          <video class="background-video dark" autoplay muted loop poster="https://webslides.tv/static/images/working.jpg">
            <source src="https://webslides.tv/static/videos/working.mp4" type="video/mp4">
          </video>
          <div class="wrap size-30">
            <p><a href="#" title="Ceara Colegial"><img class="aligncenter" src="../static/images/logos/colegia3.svg" alt="Microsoft"></a></p>
            <form id="receita_cnpj" name="receita_cnpj" action="validate.asp" method="post">
              <fieldset>
                <legend>Ceará Colegial          Sistema de Inscrição - Eventos</legend>
				                <p><label>CNPJ Escola</label>
                  <input type="text" tabindex="1" name="cnpj" id="cnpj" maxlength="14"  placeholder="CNPJ Exemplo 000.000.00/0000-00" required>
                </p>
                <p><label>Senha<span class="alignright"><a href="#" title="Forgot password?"></a></span></label>
                  <input type="password" tabindex="2" name="senha" id="senha"placeholder="Senha" required>
                </p>
                <p>
                  <button type="submit" tabindex="3" title="Login">Entrar &rsaquo;</button>
                </p>
              </fieldset>  
            </form>
            <p class="aligncenter">Sem Cadastro ? <a href="http://www.sistemainc.com.br/cearacolegial/Cadastro_Escolas/" title="Register">clique AQUI!</a></p>
			<%

If Session("allow") = False Then
  Response.Write "Voce nao esta logado."
  
Else
  Response.Write "Voce continua logado ao sistema como :<a href=""welcome.asp"">Acessar</a> " & Session("nome_fantasia") 
  Response.Write "<p><a href=""utility.asp?method=abandon"">Sair</a>"

End If
%>
          </div>
          <!-- .end .wrap -->
        </section>
		
		<script src="../static/js/webslides.js"></script>    
    <script>
      window.ws = new WebSlides();
    </script>
    
    <!-- OPTIONAL - svg-icons.js (fontastic.me - Font Awesome as svg icons) -->
    <script defer src="../static/js/svg-icons.js"></script>
  
  </body>
</html>