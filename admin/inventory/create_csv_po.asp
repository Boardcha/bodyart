<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request("po_id") = "" then
' If downloading after using the create order button on the ordering page
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM TBL_PurchaseOrders ORDER BY PurchaseOrderID DESC" 
	objCmd.Prepared = true
	Set rsGetPO_ID = objCmd.Execute()

else
	' If downloading from the purchase orders main page
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM TBL_PurchaseOrders WHERE PurchaseOrderID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("po_new_id",3,1,10,request("po_id")))
'	objCmd.Prepared = true
	Set rsGetPO_ID = objCmd.Execute()

end if

var_po_id = rsGetPO_ID.Fields.Item("PurchaseOrderID").Value

' Get most recent purchase order id #
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT (jewelry.title + ' ' + ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '')) as description, IIF(po_qty_vendor > 0, po_qty_vendor, po_qty) as po_qty, tbl_po_details.po_qty_vendor, ProductDetails.detail_code FROM jewelry  INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID INNER JOIN tbl_po_details ON ProductDetails.ProductDetailID = tbl_po_details.po_detailid WHERE tbl_po_details.po_orderid = ? AND tbl_po_details.po_qty > 0"
objCmd.Parameters.Append(objCmd.CreateParameter("po_new_id",3,1,10,var_po_id))
set rsGetItems = objCmd.Execute()

%>
<%
'***********************
' http://www.dwzone.it
' Csv Writer
' Version 1.1.4
' Start Code
'***********************
Dim dwzCsv_rs
Set dwzCsv_rs = new dwzCsvExport
dwzCsv_rs.Init
dwzCsv_rs.SetFileName "" & rsGetPO_ID.Fields.Item("Brand").Value & ".csv"
dwzCsv_rs.SetNumberOfRecord "ALL"
dwzCsv_rs.SetStartOn "ONLOAD", ""
dwzCsv_rs.SetFieldSeparator ","
dwzCsv_rs.SetFieldLabel "true"
dwzCsv_rs.SetRecordset rsGetItems
dwzCsv_rs.addItem "Qty", "po_qty", "String"
dwzCsv_rs.addItem "Item", "description", "String"
dwzCsv_rs.addItem "SKU", "detail_code", "String"
dwzCsv_rs.Execute()
'***********************
' http://www.dwzone.it
' Csv Writer
' End Code
'***********************
%>

<head>
<title>Export order</title>
</head>
<body>
</body>
</html>

<%
DataConn.Close()
%>
<!--#include virtual="dwzExport/CsvExport.asp" -->