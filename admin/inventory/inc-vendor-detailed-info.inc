<%
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT  TOP (20) d.ProductDetailID, p.title + ' ' + d.Gauge + ' ' + d.Length + ' ' + d.ProductDetail1 AS 'title', SUM(s.qty) AS 'amount_sold' " & _
		"FROM jewelry AS p INNER JOIN ProductDetails AS d ON p.ProductID = d.ProductID " & _ 
		"INNER JOIN TBL_OrderSummary AS s ON p.ProductID = s.ProductID " & _  
		"INNER JOIN sent_items AS i ON s.InvoiceID = i.ID " & _ 
		"WHERE (p.brandname = ?) AND (i.ship_code = N'paid') AND (s.item_price > 0) " & _ 
		"GROUP BY d.ProductDetailID, i.date_order_placed, p.title + ' ' + d.Gauge + ' ' + d.Length + ' ' + d.ProductDetail1 HAVING (i.date_order_placed >= GETDATE() - 90) ORDER BY amount_sold DESC" 
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,50, rsGetMainList("brand")))
	Set rsGetTopSellers = objCmd.Execute()
%>
<%
	'Get Total Sales for last 90 days
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT SUM(s.qty) As total_sales " & _
		"FROM jewelry AS p INNER JOIN ProductDetails AS d ON p.ProductID = d.ProductID " & _ 
		"INNER JOIN TBL_OrderSummary AS s ON p.ProductID = s.ProductID " & _  
		"INNER JOIN sent_items AS i ON s.InvoiceID = i.ID " & _ 
		"WHERE (p.brandname = ?) AND (i.ship_code = N'paid') AND (s.item_price > 0) " & _ 
		"AND (i.date_order_placed >= GETDATE() - 90)" 
	objCmd.Parameters.Append(objCmd.CreateParameter("value",200,1,50, rsGetMainList("brand")))
	Set rsGetSales = objCmd.Execute()
%>
<% if not rsGetTopSellers.eof then %>
	<tr>
		<td colspan="9" style="padding:0">
			<table class="admin-table">
			<thead>
				<tr style="background-color:#fff">
					<td class="font-weight-bold" style="border-top:0"><a class="font-weight-bold" href="add_company.asp#<%= rsGetMainList("brand") %>" target="_blank"><%= rsGetMainList("brand") %></a> Top 20 selling items</td>
					<td class="text-center font-weight-bold" style="border-top:0">Qty sold</td>
				</tr>
			</thead>
			<% while not rsGetTopSellers.eof %>
				<tr>
					<td>
						<%= rsGetTopSellers.Fields.Item("title").Value %></td>
					<td class="text-center">
						<%= rsGetTopSellers.Fields.Item("amount_sold").Value %>
						<div id="myProgress">
							<div style="content: '10'; position: absolute; margin: 9px 50%;"></div>
							<div id="myBar"></div>
						</div>
					</td>
				</tr>
				<%
				rsGetTopSellers.movenext()
			wend%>
			</table>
		</td>
	</tr>			
<%end if%>
