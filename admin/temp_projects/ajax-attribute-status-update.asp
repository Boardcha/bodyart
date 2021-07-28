<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("id") <> "" then

	if request.form("status") = "finished" then
		status = "attr_finished = 1"
	end if
	if request.form("status") = "review" then
		status = "attr_pass = 1"
	end if
	if request.form("status") = "discontinued" then
		status = "attr_fuckthis = 1"
	end if	
	
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE productdetails SET " + status + " WHERE productdetailID = " & request.form("id")
	objCmd.Execute()

end if

DataConn.Close()
%>