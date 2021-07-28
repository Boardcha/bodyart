<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="functions/hash_extra_key.asp"-->
<!--#include file="inc_check_matching_password.asp"-->
<%
if var_matching_password = "yes" then
%>
{  
   "matches":"yes"
}
<%	
else
%>
{  
   "matches":"no"
}	
<%

end if


DataConn.Close()
Set DataConn = Nothing
%>