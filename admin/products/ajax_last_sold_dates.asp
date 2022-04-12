<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("detailid") <> "" then
'response.write request.form("detailid")

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP (6) ORD.DetailID, Format(SNT.date_order_placed, 'yyyy-MM') as monthly, PDE.DateLastPurchased, PDE.qty AS onHand, SUM(ORD.qty) as qty_sold  FROM TBL_OrderSummary ORD INNER JOIN sent_items SNT ON ORD.InvoiceID = SNT.ID " & _
	"INNER JOIN ProductDetails PDE ON ORD.DetailID = PDE.ProductDetailID " & _
	"WHERE (SNT.ship_code = N'paid') AND (ORD.DetailID = ?) AND SNT.date_order_placed > DATEADD(MONTH, -6, GETDATE()) " & _
	"GROUP BY Format(SNT.date_order_placed, 'yyyy-MM'), ORD.DetailID, PDE.DateLastPurchased, PDE.qty " & _
	"ORDER BY monthly DESC"
	
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10,request.form("detailid")))
	Set rsGetDatesSold = objCmd.Execute()

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP (20) sent_items.ID, date_order_placed, qty, preorder, shipped FROM TBL_OrderSummary INNER JOIN sent_items ON TBL_OrderSummary.InvoiceID = sent_items.ID WHERE (sent_items.ship_code = N'paid' OR shipped = 'Cancelled') AND (TBL_OrderSummary.DetailID = ?) ORDER BY sent_items.date_order_placed DESC"
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10,request.form("detailid")))
	Set rsGetInvoicesSold = objCmd.Execute()

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT SUM(TBL_OrderSummary.qty) as 'total_on_hold' FROM TBL_OrderSummary INNER JOIN sent_items ON TBL_OrderSummary.InvoiceID = sent_items.ID WHERE (sent_items.ship_code = N'paid') AND date_order_placed > '1/1/2021' AND preorder = 1 AND TBL_OrderSummary.DetailID = ? AND  ((sent_items.shipped = N'CUSTOM ORDER IN REVIEW') OR (sent_items.shipped = N'ON ORDER'))"
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10, request.form("detailid")))
	Set rsGetTotal_PreOrderItemsOnHold = objCmd.Execute()

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP(20) sent_items.id, TBL_OrderSummary.DetailID, sent_items.shipped, TBL_OrderSummary.qty, sent_items.date_order_placed, sent_items.ship_code, sent_items.ID FROM TBL_OrderSummary INNER JOIN sent_items ON TBL_OrderSummary.InvoiceID = sent_items.ID WHERE (sent_items.ship_code = N'paid') AND date_order_placed > '1/1/2021' AND preorder = 1 AND TBL_OrderSummary.DetailID = ? AND  ((sent_items.shipped = N'CUSTOM ORDER IN REVIEW') OR (sent_items.shipped = N'ON ORDER'))"
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10, request.form("detailid")))
	Set rsGet_PreOrderItemsOnHold = objCmd.Execute()

end if
%>
<table class="table table-sm">
	<thead class="thead-dark">
	  <tr>
		<th class="py-0" scope="col">Month</th>
		<th class="py-0 text-center" scope="col">Qty sold</th>
	  </tr>
	</thead>
	<tbody>
<%
If Not rsGetDatesSold.EOF Then
	While NOT rsGetDatesSold.EOF 
	%>
	<tr>
		<td class="py-0">
			<%=rsGetDatesSold.Fields.Item("monthly").Value%>
		</td>
		<td class="py-0 text-center">
			<%= rsGetDatesSold.Fields.Item("qty_sold").Value %>
		</td>
	</tr>
	  <% 
	  total = total + rsGetDatesSold.Fields.Item("qty_sold").Value 
	  rsGetDatesSold.MoveNext()
	Wend%>
</table>
			<span class="badge badge-info" style="font-size:100%!important"><%=total%></span> sales in last 6 months
			<% if NOT rsGetTotal_PreOrderItemsOnHold.EOF then %>
			<div class="bg-dark text-light font-weight-bold p-1 my-2">Custom orders</div>
			<div>
				<span class="badge badge-info" style="font-size:100%!important"><%= rsGetTotal_PreOrderItemsOnHold("total_on_hold") %></span> on hold for custom orders
			</div>
			<% while not rsGet_PreOrderItemsOnHold.eof %>
				<a class="mr-3" href='invoice.asp?ID=<%= rsGet_PreOrderItemsOnHold("ID") %>' target='_blank' ><%= rsGet_PreOrderItemsOnHold("ID") %></a>
			<% rsGet_PreOrderItemsOnHold.movenext()
			wend
			end if %>

<%Else%>
		<div class="pt-3 py-0">
			No sales in last 6 months.
		</div>
<%
End If	
%>

<table class="table table-sm mt-4">
	<thead class="thead-dark">
	  <tr>
		<th class="py-0" scope="col">Date sold</th>
		<th class="py-0 text-center" scope="col">Qty sold</th>
	  </tr>
	</thead>
	<tbody>
<%
While NOT rsGetInvoicesSold.EOF 
%>
	  <tr>
		<td class="py-0">
			<a class="mr-3" href='invoice.asp?ID=<%=(rsGetInvoicesSold.Fields.Item("ID").Value)%>' target='_blank' ><%=FormatDateTime((rsGetInvoicesSold.Fields.Item("date_order_placed").Value), 2)%></a>
		</td>
		<td class="py-0 text-center">
			<%= rsGetInvoicesSold.Fields.Item("qty").Value %>
		</td>
	  </tr>
  <% 
  rsGetInvoicesSold.MoveNext()
Wend
%>
</tbody>
</table>

<%
set rsGetDatesSold = nothing
set rsGetInvoicesSold = nothing
DataConn.Close()
%>