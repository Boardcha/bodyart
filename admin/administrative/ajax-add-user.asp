<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO TBL_AdminUsers (name, username, AccessLevel) VALUES (?, ?, ?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,50, request.form("var_name") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,50, request.form("var_username") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,50, request.form("var_access_level") ))
	objCmd.Execute()


Set rsGetUser = nothing
DataConn.Close()
%>