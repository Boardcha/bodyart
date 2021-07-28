<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?,?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10,user_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10,request("invoiceid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,250,request("note")))
	objCmd.Execute()

DataConn.Close()
%>