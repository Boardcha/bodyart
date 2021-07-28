<%
if request.form("state") <> "" then
	session("shipping-state") = request.form("state")
else
	if request.cookies("ip-region") <> "" then
		session("shipping-state") = request.cookies("ip-region")
	else
		session("shipping-state") = ""
	end if
end if
%>