<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	'========= DELETE PURCHASE ORDER ====================================
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_PurchaseOrders SET po_hide = 1 WHERE PurchaseOrderID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("po_id",3,1,15, request("po_id")))
	objCmd.Execute()
	
%>