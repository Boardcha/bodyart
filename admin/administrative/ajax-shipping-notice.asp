<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE tbl_shipping_notice SET shipping_notice = ?, country = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("description",200,1,3000, request.form("description")))
	objCmd.Parameters.Append(objCmd.CreateParameter("country",200,1,20, request.form("country")))
	objCmd.Execute()

	response.write request.form("description")
	
DataConn.Close()
%>