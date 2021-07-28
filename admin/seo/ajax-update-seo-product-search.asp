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
	
	'response.write id
	'response.write " | value: " & column_value
	'response.write " | column: " & column_title

	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "update tbl_sitemap_searches set " & column_title & " = ? where id = ?"
'	objCmd.Parameters.Append(objCmd.CreateParameter("column",3,1,10,column_title))
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,200,column_value))
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
	objCmd.Execute()
	

Set rsGetUser = nothing
DataConn.Close()
%>