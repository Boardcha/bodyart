<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<%
response.cookies("flag-tempid") = "yes"
Response.Cookies("flag-tempid").Path = "/"
session("admin_tempcustid") = request.form("custid")

response.write session("admin_tempcustid")
%>