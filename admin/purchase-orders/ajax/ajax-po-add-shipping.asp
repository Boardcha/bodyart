<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	column_title = request.form("column")
	
	' protection to stop page from sql injection
%>
<!--#include virtual="/functions/sql_injection_prevention.asp" -->
<%
	'========= ADD SHIPPING COST TO PURCHASE ORDER ====================================
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_PurchaseOrders SET " & column_title & " = ? WHERE PurchaseOrderID = ? "
	
	if column_title = "billed_shipping_cost" then 
		'===== SET TO ALLOW CURRENCY TYPE
		objCmd.Parameters.Append(objCmd.CreateParameter("value",6,1,10, request.form("value") ))
	else 
		'===== SET TO ALLOW INTEGER TYPE
		objCmd.Parameters.Append(objCmd.CreateParameter("value",3,1,15, request.form("value") ))
	end if
	
    objCmd.Parameters.Append(objCmd.CreateParameter("po_id",3,1,15, request.form("id") ))
	objCmd.Execute()
	
%>