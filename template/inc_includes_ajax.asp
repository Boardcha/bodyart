<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/functions/sha256.asp"-->
<!--#include virtual="/functions/salt.asp"-->
<%
	if request.cookies("ip-country") <> "" then
		strcountryName = request.cookies("ip-country")
	else
		strcountryName = "US"
	end if

	if session("exchange-rate") <> "" then
	exchange_rate = session("exchange-rate")
else
	exchange_rate = 1
end if

if session("exchange-symbol") <> "" then
	exchange_symbol = session("exchange-symbol")
else
	exchange_symbol = "$"
end if

' Reset if cookie = USD
if request.cookies("currency") = "USD" OR request.cookies("currency") = "" then
        exchange_rate = 1
        exchange_symbol = "$"
        session("exchange-rate") = ""
        session("exchange-symbol") = ""
        session("exchange-currency") = ""
end if
%>
