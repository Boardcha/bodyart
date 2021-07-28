<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<%
session("exchange-rate") = request.form("rate")
session("exchange-symbol") = request.form("symbol")
session("exchange-currency") = request.form("currency")
%>
rate: <%= session("exchange-rate") %>
symbol: <%= session("exchange-symbol") %>
session currency: <%= session("exchange-currency") %>
cookie currency: <%= request.cookies("currency") %>
