<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<% response.Buffer=false
Server.ScriptTimeout=300
%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' If downloading after using the create order button on the ordering page
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM TBL_PurchaseOrders ORDER BY PurchaseOrderID DESC" 
objCmd.Prepared = true
Set rsGetPO_ID = objCmd.Execute()


if request("ID") = "" then
	var_po_id = rsGetPO_ID.Fields.Item("PurchaseOrderID").Value
else
	var_po_id = request("ID")
end if


' Get most recent purchase order id #
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT (jewelry.title + ' ' + ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '')) as description, tbl_po_details.po_qty, tbl_po_details.po_qty_vendor, ProductDetails.detail_code, ProductDetails.ProductDetailID, ProductDetails.BinNumber_Detail, ProductDetails.location, po_invoice_number,  TBL_Barcodes_SortOrder.ID_Description, TBL_PurchaseOrders.Brand, (SELECT PreOrder_Desc FROM TBL_OrderSummary WHERE OrderDetailID = tbl_po_details.po_invoice_order_detailid) AS PreOrder_Desc  FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID INNER JOIN tbl_po_details ON ProductDetails.ProductDetailID = tbl_po_details.po_detailid INNER JOIN TBL_PurchaseOrders ON tbl_po_details.po_orderid = TBL_PurchaseOrders.PurchaseOrderID LEFT OUTER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE tbl_po_details.po_orderid = ? AND tbl_po_details.po_qty > 0 ORDER BY jewelry ASC, po_invoice_number ASC, 'description' ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("po_new_id",3,1,10,var_po_id))
set rsGetItems = objCmd.Execute()
%>
<html class="simple-table">
<head>
	<title>View Order</title>
<style>
.simple-table body {background-color: #fff; color: #000; font-family: sans-serif,verdana,arial; font-size: 1em; padding: 20px;}
.simple-table table {table-layout: fixed;}
.simple-table table, .simple-table tr, .simple-table td {margin: 0; padding: 0; border: none;}
.simple-table td, .simple-table th {padding: 10px 30px 10px 10px; margin: 0; border: 1px solid black; font-family: sans-serif,verdana,arial; color: #000; text-align: left; font-size: 1em;}
.simple-table thead {background-color: #BBBBBB;}
</style>
</head>
<body>
<strong>SHIP TO:</strong><br />
Bodyartforms<br />
1966 S. Austin Ave.<br />
Georgetown, TX  78626<br />
(512) 809-3332<br /><br />
<table class="simple-table" style="border-collapse:collapse;">
<thead>
  <tr>
    <th>Invoice #</th>
	<th>Qty</th>
	<th>SKU / Code</th>
    <th>Item</th>
	<% if rsGetItems.Fields.Item("brand").Value = "Etsy" then %>
	<th>Location</th>
	<% end if%>
  </tr>
</thead>
<%
While NOT rsGetItems.EOF 
%>
  <tr>
	  <td align="left">
		<% if  rsGetItems("po_invoice_number") > 0 then %>
			<%= rsGetItems("po_invoice_number") %>
		<% end if %>
	</td>
    <td align="center">
	<%If rsGetItems.Fields.Item("po_qty_vendor").Value > 0 Then
		Response.Write rsGetItems.Fields.Item("po_qty_vendor").Value
	Else
		Response.Write rsGetItems.Fields.Item("po_qty").Value
	End If%>
	</td>
	<td align="left">
		<%= rsGetItems.Fields.Item("detail_code").Value %>
    </td>
    <td align="left">
		<% if  rsGetItems("po_invoice_number") > 0 then %>
			<%=Replace(rsGetItems("description"), "CUSTOM ORDER ", "")%>
			<br>
			 Specs: <%= rsGetItems("PreOrder_Desc") %>
		<% else %>
			<%= rsGetItems("description") %>
		<% end if %>
	</td>
	<% if rsGetItems.Fields.Item("brand").Value = "Etsy" then %>
    <td align="left">
		<% if rsGetItems.Fields.Item("BinNumber_Detail").Value <> 0 then %> 
		BIN # <%= rsGetItems.Fields.Item("BinNumber_Detail").Value %> &nbsp;&nbsp;&nbsp; <%= rsGetItems.Fields.Item("ProductDetailID").Value %>
		<% else %>
		<%= rsGetItems.Fields.Item("ID_Description").Value %>&nbsp;<%= rsGetItems.Fields.Item("location").Value %>
		<% end if %>
    </td>	
	<% end if %>
  </tr>
<%
  rsGetItems.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
DataConn.Close()
%>
