<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Buffer = True%>
<!--#INCLUDE FILE="config.inc"-->
<!--#INCLUDE FILE="level1.inc"-->
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
  <section class="bg-secondary">
  <span class="background dark" style="background-image:url('http://sistemainc.com.br/cearacolegial/fundos/photo-1453946610176-6be21147c400.jpg')"></span>
          <div class="wrap">
		   <h3>Sistema de Inscrição - Eventos</h3>
            <h2>Olá,<%Response.Write(Session("nome_fantasia"))%></h2>
            <!-- li>a? Add blink = <ul class="flexblock metrics blink">-->
            <ul class="flexblock metrics border"><p><a href=""</a></p>
              <li> 
			  <span>
                   <!-- <svg class="fa-ban"> -->
				  <svg class="fa-trophy">
                   <use xlink:href="#fa-trophy"></use><p><a href="../Entradas_escolas/eventos_new/index.asp"</a></p>
					<use xlink:href="#fa-trophy"></use><p><a href="#"</a></p>
                  </svg>
                </span>
                INSCRIÇÃO 2018<a href="../Entradas_escolas/eventos_new/index.asp"></a>				
              </li>
              <li>
                <span>
                  <svg class="fa-user-plus">
                    <use xlink:href="#fa-user-plus"></use><p><a href="Cad_alunos/df_inclusao.asp"</a></p>
                  </svg>
                </span>
                Cadastro de Atletas<a href="Cad_alunos/df_inclusao.asp"></a>
              </li>
              <li>
                <span>
                  <svg class="fa-search-plus">
                    <use xlink:href="#fa-search-plus"></use><p><a href="Cad_alunos/df_consulta.asp"</a></p>
                  </svg>
                </span>
                Consultar Atletas e imprimir Fichas<a href="Cad_alunos/df_consulta.asp"></a>
              </li>
              <li>
                <span>
                  <svg class="fa-refresh">
                    <use xlink:href="#fa-refresh"></use><p><a href="../Entradas_escolas/eventos_new/pesquisa_eventos/df_consulta.asp"</a></p>
                  </svg>
                </span>
                Consultar Evento e imprimir Fichas<a href="../Entradas_escolas/eventos_new/pesquisa_eventos/df_consulta.asp"></a>
              </li>
              
               <li>
               <span>
                  <svg class="fa-user-secret">
                    <use xlink:href="#fa-user-secret"></use><p></p>
                  </svg>
                </span>
               CNPJ:<%=Session("admin")%>
                 <br>
               IP: <%=Request.ServerVariables("REMOTE_ADDR")%>           
              </li>
              <li> 
              <span>
                  <svg class="fa-certificate">
                    <use xlink:href="#fa-certificate"></use><p><a href="http://cearacolegial1.tempsite.ws/certificado/"</a></p>
                  </svg>
                </span>
                Certificado de Participação<p><a href="http://cearacolegial1.tempsite.ws/certificado/"</a></p>              
              </li>
              <li>
              </li>
              <li>
                <span>
                  <svg class="fa-sign-out">
                    <use xlink:href="#fa-sign-out"></use><p><a href="logout.asp"</a></p>
                  </svg>
                </span>
                Sair Sistema<p><a href="logout.asp"</a></p>
              </li>
              </ul>
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