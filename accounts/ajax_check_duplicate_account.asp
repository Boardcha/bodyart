<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/accounts/inc_check_duplicate_account.asp"-->
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
<%

end if


DataConn.Close()
Set DataConn = Nothing
%>