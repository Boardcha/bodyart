<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' ====== UPDATE THE TIME SCANNED FIELD =======
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE TBL_OrderSummary SET TimesScanned = TimesScanned + 1 WHERE OrderDetailID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("OrderDetailID",3,1,15, request.form("OrderDetailID") ))
objCmd.Execute  
%>
