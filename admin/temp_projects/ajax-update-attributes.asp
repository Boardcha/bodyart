<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

if request.form("id") <> "" then
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE productdetails SET " + request.form("column") + " = '"  + request.form("value") + "'  WHERE productdetailID = " & request.form("id")
	objCmd.Execute()

end if

DataConn.Close()
%>