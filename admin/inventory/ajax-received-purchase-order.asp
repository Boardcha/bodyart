<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	'========= UPDATE PURCHASE ORDER TO RECEIVED ====================================
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_PurchaseOrders SET Received = 'Y', DateReceived = CAST( GETDATE() AS Date ) WHERE PurchaseOrderID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("po_id",3,1,15, request("po_id")))
	objCmd.Execute()
	
%>