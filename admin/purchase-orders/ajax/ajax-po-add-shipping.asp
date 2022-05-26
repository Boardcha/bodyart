<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	'========= ADD SHIPPING COST TO PURCHASE ORDER ====================================
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_PurchaseOrders SET billed_shipping_cost = ? WHERE PurchaseOrderID = ? "
	objCmd.Parameters.Append(objCmd.CreateParameter("value",6,1,10, request("value") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("po_id",3,1,15, request("id") ))
	objCmd.Execute()
	
%>