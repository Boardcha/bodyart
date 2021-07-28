<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET over_150 = 0 WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, request.form("invoice_id") ))
	objCmd.Execute()

DataConn.Close()
%>