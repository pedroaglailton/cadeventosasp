<%@Language=VBScript%>
<%Response.Buffer=True%>

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
  Response.Write "Voce não tem autorização para  esta página!"
  Response.Write "<p><a href=""index.asp"">Voltar</a>"
End Sub

Sub Expired()
  Response.Write "Sistema de registro de encerrado"
  Response.Write "<p>Contate a Ceara Colegial "
  Session.Abandon
End Sub

Sub Abandon()
  Session.Abandon
  Response.Redirect "index.asp"
End Sub
%>
