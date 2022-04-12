<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	column_title = request.form("column")
	
	' protection to stop page from sql injection
	if Len(column_title) > 25 then
		response.end
	end if

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "update TBL_OrderSummary set " & column_title & " = ? where OrderDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,225, server.htmlencode(request.form("value")) ))
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10, request.form("id") ))
	objCmd.Execute()

DataConn.Close()
%>