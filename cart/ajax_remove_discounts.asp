{
<%
' Remove STORE CREDIT ----------------
if request.form("remove_type") = "store-credit" then
	session("usecredit") = ""
	session("storeCredit_used") = 0
%>
	"remove_type":"store-credit"
<%
end if
%>
}