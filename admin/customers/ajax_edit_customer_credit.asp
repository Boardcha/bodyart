<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET credits = ? WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("amount",6,1,10,request("amount")))
	objCmd.Parameters.Append(objCmd.CreateParameter("custid",3,1,10,request("custid")))
	objCmd.Execute()

DataConn.Close()
%>