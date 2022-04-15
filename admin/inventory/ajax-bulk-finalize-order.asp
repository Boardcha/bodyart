<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' Insert new purchase order into table	
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "INSERT INTO TBL_PurchaseOrders (po_internal_bulk_pull) VALUES (1)"
objCmd.Execute()

' Get most recent purchase order id #
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TOP 1 PurchaseOrderID FROM TBL_PurchaseOrders ORDER BY PurchaseOrderID DESC"
set rsGetPO_ID = objCmd.Execute()

'====== IF THE PURCHASE ORDER NEEDS TO BE REVIEWED BY A MANAGER FLAG THE ORDER =================
if request.form("var_needs_review") = "yes" then
    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "UPDATE TBL_PurchaseOrders SET po_needs_review = 1 WHERE PurchaseOrderID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("po_new_id",3,1,20,rsGetPO_ID.Fields.Item("PurchaseOrderID").Value))
    objCmd.Execute()
end if



' Replace all temp ID's ordered with new purchase order #	
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE tbl_po_details SET po_orderid = ?, po_temp_id = 0 WHERE po_temp_id = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("po_new_id",3,1,20,rsGetPO_ID.Fields.Item("PurchaseOrderID").Value))
objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,20, request.Cookies("bulk-po-id") ))
objCmd.Execute()

Response.Cookies("bulk-po-id") = ""
Response.Cookies("bulk-po-id").Expires = DATE - 1

Set rsGetUser = nothing
DataConn.Close()
%>
