<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("id") <> "" then

	if request.form("status") = "delete" then
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "DELETE FROM customers WHERE customer_ID = " & request.form("id")
		objCmd.Execute()
	end if
	
end if

DataConn.Close()
%>