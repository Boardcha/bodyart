<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_AdminUsers SET archived = 1, password_hashed = '', salt = '', AccessLevel = '' where ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10, request.form("user_id") ))
	objCmd.Execute()
	

Set rsGetUser = nothing
DataConn.Close()
%>