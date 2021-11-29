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
objCmd.CommandText = "SELECT (jewelry.title + ' ' + ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '')) as description, tbl_po_details.po_qty, tbl_po_details.po_qty_vendor, ProductDetails.detail_code, ProductDetails.ProductDetailID, ProductDetails.BinNumber_Detail, ProductDetails.location, TBL_Barcodes_SortOrder.ID_Description, dbo.TBL_PurchaseOrders.Brand FROM dbo.jewelry INNER JOIN dbo.ProductDetails ON dbo.jewelry.ProductID = dbo.ProductDetails.ProductID INNER JOIN dbo.tbl_po_details ON dbo.ProductDetails.ProductDetailID = dbo.tbl_po_details.po_detailid INNER JOIN dbo.TBL_PurchaseOrders ON dbo.tbl_po_details.po_orderid = dbo.TBL_PurchaseOrders.PurchaseOrderID LEFT OUTER JOIN dbo.TBL_Barcodes_SortOrder ON dbo.ProductDetails.DetailCode = dbo.TBL_Barcodes_SortOrder.ID_Number WHERE tbl_po_details.po_orderid = ? AND tbl_po_details.po_qty > 0 ORDER BY 'description' ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("po_new_id",3,1,10,var_po_id))
set rsGetItems = objCmd.Execute()
%>
<html class="simple-table">
<head>
    <title>Bulk order</title>
<style>
.simple-table body {background-color: #fff; color: #000; font-family: sans-serif,verdana,arial; font-size: 1em; padding: 20px;}
.simple-table table {table-layout: fixed;}
.simple-table table, .simple-table tr, .simple-table td {margin: 0; padding: 0; border: none;}
.simple-table td, .simple-table th {padding: 10px 30px 10px 10px; margin: 0; border: 1px solid black; font-family: sans-serif,verdana,arial; color: #000; text-align: left; font-size: 1em;}
.simple-table thead {background-color: #BBBBBB;}
</style>
</head>
<body>
<table class="simple-table" style="border-collapse:collapse;">
<thead>
  <tr>
    <th>Qty</th>
    <th>Location</th>
	<th>Item</th>
  </tr>
</thead>
<%
While NOT rsGetItems.EOF 
%>
  <tr>
    <td align="center">
	<%= rsGetItems("po_qty") %>
	</td>
    <td align="left">
		<%= rsGetItems.Fields.Item("description").Value %>
    </td>
    <td align="left">
		<% if rsGetItems.Fields.Item("BinNumber_Detail").Value <> 0 then %> 
		BIN # <%= rsGetItems.Fields.Item("BinNumber_Detail").Value %> &nbsp;&nbsp;&nbsp; <%= rsGetItems.Fields.Item("ProductDetailID").Value %>
		<% else %>
		<%= rsGetItems.Fields.Item("ID_Description").Value %>&nbsp;<%= rsGetItems.Fields.Item("location").Value %>
		<% end if %>
    </td>	
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
