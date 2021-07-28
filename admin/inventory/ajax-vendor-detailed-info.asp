<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<h1>Under construction</h1>
<%
var_brand = request.querystring("brand")

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT  TOP (30) d.ProductDetailID, p.title + ' ' + d.Gauge + ' ' + d.Length + ' ' + d.ProductDetail1 AS 'title', SUM(s.qty) AS 'amount_sold' FROM jewelry AS p INNER JOIN ProductDetails AS d ON p.ProductID = d.ProductID INNER JOIN TBL_OrderSummary AS s ON p.ProductID = s.ProductID INNER JOIN sent_items AS i ON s.InvoiceID = i.ID WHERE (p.brandname = ?) AND (i.ship_code = N'paid') AND (s.item_price > 0) GROUP BY d.ProductDetailID, i.date_order_placed, p.title + ' ' + d.Gauge + ' ' + d.Length + ' ' + d.ProductDetail1 HAVING  (i.date_order_placed >= GETDATE() - 90) ORDER BY i.date_order_placed DESC" 
objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,50, var_brand))
Set rsGetTopSellers = objCmd.Execute()
%>
TEST DATA .... IT'S NOT RIGHT YET
<% if not rsGetTopSellers.eof then %>
<table class="admin-table">
<thead>
<tr>
	<td colspan="2">Top 30 selling items</td>
</tr>
<tr>
	<td>Item</td>
	<td>Qty sold</td>
</tr>
</thead>
<% while not rsGetTopSellers.eof %>
<tr>
<td>
<%= rsGetTopSellers.Fields.Item("title").Value %></td>
<td>
<%= rsGetTopSellers.Fields.Item("amount_sold").Value %>
</td>
</tr>
<%
rsGetTopSellers.movenext()
wend
end if
%>
<%
DataConn.Close()
%>