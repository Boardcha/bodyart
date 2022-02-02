<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
invoiceid = request.form("id")

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM sent_items WHERE ID = ? AND ship_code = 'paid' AND (shipped = 'Pending...' OR shipped = 'Pending shipment' OR shipped = 'Review' OR shipped = 'CUSTOM ORDER IN REVIEW')"
	objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,invoiceid))
	Set rsGetOrder = objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT sent_items.shipping_rate - sent_items.total_preferred_discount - sent_items.total_coupon_discount - sent_items.total_free_credits - sent_items.total_returns + sent_items.total_sales_tax + sent_items.total_store_credit + sent_items.total_gift_cert AS total_discount_taxes, SUM(TBL_OrderSummary.qty * TBL_OrderSummary.item_price) AS subtotal FROM sent_items INNER JOIN TBL_OrderSummary ON sent_items.ID = TBL_OrderSummary.InvoiceID WHERE (sent_items.ID = ?) GROUP BY sent_items.shipping_rate - sent_items.total_preferred_discount - sent_items.total_coupon_discount - sent_items.total_free_credits - sent_items.total_returns + sent_items.total_sales_tax + sent_items.total_store_credit + sent_items.total_gift_cert"
	objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,invoiceid))
	Set rsGetOrderTotal = objCmd.Execute()
	
	if not rsGetOrderTotal.eof then
		var_order_total = rsGetOrderTotal.Fields.Item("subtotal").Value + rsGetOrderTotal.Fields.Item("total_discount_taxes").Value
	end if

if not rsGetOrder.eof then	

if CLng(CustID_Cookie) = CLng(rsGetOrder.Fields.Item("customer_ID").Value) then
%>

	If you do not want a store credit, please contact <a class="font-weight-bold text-info" href="/contact.asp">customer service</a>.
	<p>By cancelling below, you'll be issued a store credit for the full amount of your order and it will be available for use on your account immediately after confirming cancellation.
		</p>
	
	<div class="alert alert-success py-1 my-2">
	<span class="font-weight-bold"><%= FormatCurrency(var_order_total, -1, -2, -0, -2) %></span> will be applied to your <span class="font-weight-bold">store credit</span>
	</div>
<% else ' ' CustID_Cookie = order customerID
%>
Unauthorized access
<%
	end if ' CustID_Cookie = order customerID
else '==== order no longer eligible to be cancelled
%>
<div class="alert alert-danger py-1 my-2">
	Order has been shipped. Please <a href="/contact.asp">contact customer service</a>
	</div>
<%
end if ' not rsGetOrder.eof
DataConn.Close()
%>
