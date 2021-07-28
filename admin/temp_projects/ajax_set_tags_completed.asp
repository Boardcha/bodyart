<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("id") <> "" then

	'Set item as tagged complete
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE jewelry SET tags_completed = 1 WHERE ProductID = " & request.form("id")
	objCmd.Execute()

end if

DataConn.Close()
%>