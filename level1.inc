<% 
If session("allow") = False Then Response.Redirect "index.asp"
If session("clearance") < 1 Then Response.Redirect "utility.asp?method=unauthorized"
%>
