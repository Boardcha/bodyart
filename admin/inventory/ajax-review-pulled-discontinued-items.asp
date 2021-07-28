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
	objCmd.CommandText = "UPDATE ProductDetails SET qty = ? where ProductDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,15, column_value ))
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,15, id))
	objCmd.Execute()
		

Set rsGetUser = nothing
DataConn.Close()
%>