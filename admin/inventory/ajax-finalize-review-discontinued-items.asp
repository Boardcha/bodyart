<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE jewelry SET to_be_pulled = 0 where ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,15, request.form("productid") ))
	objCmd.Execute()
		

Set rsGetUser = nothing
DataConn.Close()
%>