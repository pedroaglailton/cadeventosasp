<% 
If session("allow") = False Then Response.Redirect "http://www.sistemainc.com.br/cearacolegial/sistema/index.asp"
If session("clearance") < 1 Then Response.Redirect "utility.asp?method=unauthorized"
%>
