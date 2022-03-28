<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' ====== UPDATE THE TIME SCANNED FIELD =======
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE TBL_OrderSummary SET TimesScanned = 0 WHERE InvoiceID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, request.form("invoiceid") ))
objCmd.Execute  
%>
