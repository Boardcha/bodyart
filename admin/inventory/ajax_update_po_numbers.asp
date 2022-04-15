<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	' protection to stop page from sql injection
%>
<!--#include file="../../functions/sql_injection_prevention.asp" -->
<%
	'=== GET TEMP PO ID START DATE SO THAT WE CAN INSERT IT INTO THE NEWLY CREATED PURCHASE ORDER
	' Retrieve the newest temp PO # to use for saving order details
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT po_time_started FROM tbl_po_temp_ids WHERE po_temp_id = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,10,request("tempid")))
	Set rsGetStartDate = objCmd.Execute()

		if NOT rsGetStartDate.EOF then
			po_start_date = rsGetStartDate.Fields.Item("po_time_started").Value
		end if

	' Insert new purchase order into table	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO TBL_PurchaseOrders (Brand, po_total, po_time_started, months_to_restock) VALUES (?, ?, ?, ?)"
	objCmd.Parameters.Append objCmd.CreateParameter("brand",200,1,50, request("brand"))
	objCmd.Parameters.Append(objCmd.CreateParameter("po_total",6,1,10,request("pototal")))
	objCmd.Parameters.Append(objCmd.CreateParameter("po_time_started",200,1,30, Cstr(po_start_date)))
	objCmd.Parameters.Append(objCmd.CreateParameter("for_how_many_months",3,1,2, request("for_how_many_months")))
	objCmd.Execute()

	' Get most recent purchase order id #
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP 1 PurchaseOrderID FROM TBL_PurchaseOrders ORDER BY PurchaseOrderID DESC"
	set rsGetPO_ID = objCmd.Execute()

	' Replace all temp ID's ordered with new purchase order #	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE tbl_po_details SET po_orderid = ?, po_temp_id = 0 WHERE po_temp_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("po_new_id",3,1,10,rsGetPO_ID.Fields.Item("PurchaseOrderID").Value))
	objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,10,request("tempid")))
	objCmd.Execute()

	' Deduct inventory if it's an Etsy PO
	if request("brand") = "Etsy" then
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE vw_etsy_purchase_order_deduct SET qty = qty - po_qty WHERE PurchaseOrderID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("po_new_id",3,1,10,rsGetPO_ID.Fields.Item("PurchaseOrderID").Value))
		objCmd.Execute()	
	end if

DataConn.Close()

	' Delete cookie that holds the temp ID # so that it can start a new order from scratch
	response.cookies("brandname = '" & request("brand") & "'") = ""
	response.cookies("brandname = '" & request("brand") & "'").Expires = DateAdd("d",-1,now())
	response.cookies("po-filter-status") = ""
	response.cookies("po-filter-active") = ""
	response.cookies("po-filter-qty") = ""
%>