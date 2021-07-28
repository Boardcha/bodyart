<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'------ SET TIMESTAMP FOR ORDER COMPLETED SCAN

' ====== UPDATE ORDER WITH PULLER INFO, DATE STARTED
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE sent_items SET time_pulled_finished = GETDATE() WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, request.form("invoiceid") ))
objCmd.Execute  
%>
