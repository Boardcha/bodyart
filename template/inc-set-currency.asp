<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<%
if request.form("currency") <> "" then
	response.cookies("currency") = request.form("currency")
	
	if request.form("currency") = "GBP" then
		response.cookies("currency-symbol") = "£"
	elseif request.form("currency") = "EUR" then
		response.cookies("currency-symbol") = "€"
	elseif request.form("currency") = "JPY" then
		response.cookies("currency-symbol") = "¥"
	elseif request.form("currency") = "DKK" then
		response.cookies("currency-symbol") = "kr"
	else
		response.cookies("currency-symbol") = "$"
	end if

end if
%>