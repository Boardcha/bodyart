<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	column_title = request.form("column")
	
	' protection to stop page from sql injection
%>
<!--#include file="../../functions/sql_injection_prevention.asp" -->
<%
	
	column_value = request.form("value")
	id = request.form("id")
	friendly_name = request.form("friendly_name")

	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "update TBL_AdminUsers set " & column_title & " = ? where ID = ?"
'	objCmd.Parameters.Append(objCmd.CreateParameter("column",3,1,10,column_title))
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,50,column_value))
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
	objCmd.Execute()
	

Set rsGetUser = nothing
DataConn.Close()
%>