<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% 

' response.write request.form("e-mail") & "<br/>" & request.form("password") & "<br/>" &request.form("shipping-first") & "<br/>" & request.form("shipping-last")

%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="connections/authnet.asp"-->
<!--#include file="../functions/token.asp"-->
<!--#include file="../functions/hash_extra_key.asp"-->
<!--#include file="inc_check_duplicate_account.asp"-->
<%
if var_duplicate_account = "yes" then
%>
{  
   "duplicate":"yes"
}
<%	
else
%>
{  
   "duplicate":"no"
}	
<!--#include file="inc_create_account.asp"-->
<%

end if


DataConn.Close()
Set DataConn = Nothing
%>