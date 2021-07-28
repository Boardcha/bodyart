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
	int_string = request.form("int_string")
	
'	response.write "Friendly: " & friendly_name & " Column" & column_title

	if int_string = "string" then
	'	update text fields
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "update customers set " & column_title & " = ? where customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,50,column_value))
		objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
		objCmd.Execute()
	
	elseif int_string = "integer" then ' integer updates only

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "update customers set " & column_title & " = ? where customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("value",3,1,10,column_value))
		objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))
		objCmd.Execute()
		
	elseif int_string = "money" then ' money updates only

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "update customers set " & column_title & " = ? where customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("value",6,1,10,column_value))
		objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,id))

		objCmd.Execute()	
	end if
	

	

Set rsGetUser = nothing
DataConn.Close()
%>