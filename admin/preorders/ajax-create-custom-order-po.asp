<%@LANGUAGE="VBSCRIPT" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT InvoiceID, ProductID, DetailID, qty, PreOrder_Desc, detail_code, title, ProductDetail1, OrderDetailID, Gauge, Length, OrderDetailID FROM QRY_OrderDetails WHERE customorder = 'yes' AND (shipped = 'CUSTOM ORDER IN REVIEW' or shipped = 'ON ORDER') AND item_ordered = 0 AND brandname = ? ORDER BY jewelry, InvoiceID ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("brandname",200,1,100, request.form("brandname") ))
Set rsGetPreorders = objCmd.Execute()

	'====== CREATE A NEW EMPTY PURCHASE ORDER
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO TBL_PurchaseOrders (DateOrdered, Brand, po_type) VALUES ('"& date() &"', '"& request.form("brandname") & "', 'Custom Orders')" 
	objCmd.Execute()	

	'====== FIND THE LAST PURCHASE ORDER ID IN TABLE TO WRITE THE ITEMS TO
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP 1 PurchaseOrderID FROM TBL_PurchaseOrders ORDER BY PurchaseOrderID DESC"  
	Set rsGetLastPO = objCmd.Execute


    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn

WHILE NOT rsGetPreorders.EOF

    objCmd.CommandText = "UPDATE TBL_OrderSummary SET item_ordered = 1, item_ordered_date = GETDATE() WHERE OrderDetailID = " & rsGetPreorders("OrderDetailID") 
    objCmd.Execute()

    objCmd.CommandText = "UPDATE sent_items SET shipped = 'ON ORDER', date_sent = GETDATE() WHERE ID = " & rsGetPreorders("InvoiceID")  
    objCmd.Execute()

    '====== INSERT ITEM INTO PURCHASE ORDER DETAILS TABLE
    objCmd.CommandText = "INSERT INTO tbl_po_details (po_qty, po_orderid, po_detailid, po_invoice_number, po_invoice_order_detailid, po_confirmed) VALUES (" & rsGetPreorders("qty") & ", " & rsGetLastPO("PurchaseOrderID") & ", " & rsGetPreorders("DetailID") & ", " & rsGetPreorders("InvoiceID") & ", " & rsGetPreorders("OrderDetailID") & ", 1)"
    objCmd.Execute()

rsGetPreorders.MoveNext()
WEND
%>

{
	"purchase_order_id" : "<%= rsGetLastPO("PurchaseOrderID") %>"
}

<%
rsGetPreorders.Close()
Set rsGetPreorders = Nothing
Set objCmd = Nothing
%>
