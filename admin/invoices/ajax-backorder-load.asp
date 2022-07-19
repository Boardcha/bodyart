<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT customer_ID, total_preferred_discount, total_coupon_discount, coupon_code FROM sent_items WHERE ID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("string_id",3,1,12,request.form("invoice")))
Set rsGetOrder = objCmd.Execute()

if not rsGetOrder.eof then
	custID = rsGetOrder.Fields.Item("customer_ID").Value
	var_coupon_discount = rsGetOrder.Fields.Item("total_coupon_discount").Value
	var_preferred_discount = rsGetOrder.Fields.Item("total_preferred_discount").Value
	var_coupon_code = rsGetOrder.Fields.Item("coupon_code").Value
end if

'if rsGetOrder.Fields.Item("coupon_code").Value <> "" then
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT DiscountPercent FROM TBLDiscounts WHERE DiscountCode = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("coupon_code",200,1,50,rsGetOrder.Fields.Item("coupon_code").Value))
	Set rsGetCouponDiscount = objCmd.Execute()
'end if


if var_preferred_discount <> 0 or var_coupon_discount <> 0 or var_coupon_code <> "" then
	if NOT rsGetCouponDiscount.eof then	
	'	var_price = FormatNumber((var_price - ((rsGetCouponDiscount.Fields.Item("DiscountPercent").Value / 100) * var_price)) * request.form("qty"), -1, -2, -0, -2)
		var_discount = rsGetCouponDiscount.Fields.Item("DiscountPercent").Value
	else
		var_discount = 0
	end if
end if		

' Calculate Refund amount for the item
var_invoice_number = request.form("invoice")
orderDetailID = request.form("item")
ProductDetailID = request.form("detailid")

set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT sent_items.ID, sent_items.coupon_code, sent_items.combined_tax_rate, TBL_OrderSummary.ErrorReportDate, TBL_OrderSummary.ErrorDescription,  sent_items.ship_code, TBL_OrderSummary.qty, ProductDetails.qty AS 'qty_instock', TBL_OrderSummary.item_price, ProductDetails.ProductDetail1, ProductDetails.location, ProductDetails.Gauge, ProductDetails.Length, jewelry.title, ProductDetails.ProductDetailID, ProductDetails.BinNumber_Detail, TBL_OrderSummary.OrderDetailID, TBL_OrderSummary.ProductID, TBL_OrderSummary.item_problem, TBL_OrderSummary.ErrorQtyMissing,  (jewelry.title + ' ' + ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '')) as description FROM sent_items INNER JOIN TBL_OrderSummary ON sent_items.ID = TBL_OrderSummary.InvoiceID INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID WHERE TBL_OrderSummary.backorder = 1 AND ID = ? AND TBL_OrderSummary.OrderDetailID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, var_invoice_number))
objCmd.Parameters.Append(objCmd.CreateParameter("orderDetailID",3,1,12, orderDetailID))

set rsGetItem = Server.CreateObject("ADODB.Recordset")
rsGetItem.CursorLocation = 3 'adUseClient
rsGetItem.Open objCmd

If NOT rsGetItem.EOF Then
	'==============  GET COUPON DISCOUNT / IF ANY ============================================
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT DiscountPercent FROM TBLDiscounts WHERE DiscountCode = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("coupon_code",200,1,50,rsGetItem.Fields.Item("coupon_code").Value))
	Set rsGetCouponDiscount = objCmd.Execute()
End If

If Not rsGetItem.EOF Then

	If NOT rsGetCouponDiscount.eof then
		var_item_price = FormatNumber((rsGetItem.Fields.Item("item_price").Value - ((rsGetCouponDiscount.Fields.Item("DiscountPercent").Value / 100) * rsGetItem.Fields.Item("item_price").Value)) * rsGetItem.Fields.Item("qty").Value, -1, -2, -0, -2)                        
	Else
		var_item_price = FormatNumber(rsGetItem.Fields.Item("item_price").Value * rsGetItem.Fields.Item("qty").Value, -1, -2, -0, -2)
	End if

	' Add on tax to refund 
	If rsGetItem.Fields.Item("combined_tax_rate").Value > 0 then
		var_item_price = var_item_price + (var_item_price * rsGetItem.Fields.Item("combined_tax_rate").Value)
	End if

	var_item_refund = FormatNumber(Ccur(var_item_refund) + ccur(var_item_price), -1, -2, -0, -2)
End If

If var_item_refund > 0 then
	'Add shipping price to refund amount if it is the only item in the order Or when all the items are backordered in the order (free items are excluded)
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT shipping_rate, retail_delivery_fee FROM sent_items WHERE ID = ? AND ( " & _
		"SELECT TOP 1 ORS.InvoiceID FROM TBL_OrderSummary ORS " & _
		"LEFT JOIN sent_items SNT ON SNT.ID = ORS.InvoiceID " & _
		"INNER JOIN ProductDetails DET ON DET.ProductDetailID = ORS.DetailID " & _
		"WHERE ORS.InvoiceID = ? AND ORS.DetailID <> ? AND ORS.backorder <> 1 AND (DET.free = 0 AND DET.ProductID not in(1464, 1483, 1649, 2991, 3086, 3587, 3611, 3803, 3926, 3928, 4287))) is null"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, var_invoice_number))
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid2",3,1,12, var_invoice_number))
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,15, ProductDetailID))
	Set rsGetShippingRate = objCmd.Execute()
	Response.Write var_item_refund & "<br>"
	If Not rsGetShippingRate.EOF Then
		var_item_refund = FormatNumber(Ccur(var_item_refund) + Ccur(rsGetShippingRate("shipping_rate")) + Ccur(rsGetShippingRate("retail_delivery_fee")), -1, -2, -0, -2)
	End If
End If	

%>

<div class="bo-message"></div>

<input type="hidden" id="qty_<%= request.form("item") %>" value="<%= request.form("qty") %>">
<input type="hidden" id="detailid_<%= request.form("item") %>" value="<%= request.form("detailid") %>">
<input type="hidden" id="price_<%= request.form("item") %>" value="<%= request.form("price") %>">
<input type="hidden" id="origprice_<%= request.form("item") %>" value="<%= request.form("origprice") %>">

<div class="container backorders">
	<div class="row">
<div class="col">
<% if cint(request.form("qty_instock")) >= cint(request.form("qty")) then %>
	<button class="btn btn-sm btn-secondary my-1 btn_bo" data-agenda="ship-one" data-item="<%= request.form("item") %>">Ship out item</button>
	<button class="btn btn-sm btn-secondary my-1 btn_bo" data-agenda="reship" data-item="<%= request.form("item") %>">Reship current order</button>
	<% else %>
	<div class="alert alert-danger">Not enough in stock to reship</div><br/>
<% end if %>
<button class="btn btn-sm btn-secondary my-1 btn_bo" data-agenda="clear" data-item="<%= request.form("item") %>">Clear backorder</button>

<div class="h6 mt-4">Refund item only $<%= var_item_refund %></div>
<button class="btn btn-sm btn-secondary btn_bo" data-agenda="item-refund" data-item="<%= request.form("item") %>">Refund</button>

<% if custID <> 0 then %>
	<button class="btn btn-sm btn-secondary btn_bo" data-agenda="item-storecredit" data-item="<%= request.form("item") %>">Store credit</button>
<% end if %>
<div class="h6 mt-4">Cancel entire order $<%= request.form("total") %></div>
<button class="btn btn-sm btn-secondary btn_bo" data-agenda="cancel-refund" data-item="<%= request.form("item") %>">Refund</button>
<% if custID <> 0 then %>
	<button class="btn btn-sm btn-secondary btn_bo" data-agenda="cancel-storecredit" data-item="<%= request.form("item") %>">Store credit</button>
<% end if %>
</div>

<div class="col">
<h6>Exchange</h6>
<div class="float-left">
		<div class="form-group">
	<input class="form-control form-control-sm" type="text" id="bo-exchange-product" placeholder="Product #">
		</div>
	<div id="bo-exchange-form" style="display:none">
		<form>
				<div class="form-group">
			<label>Qty</label>
			<input  class="form-control form-control-sm" type="text" id="bo-exchange-qty" value="1">
			</div>
			<div class="form-group">
			<label>Detail #</label>
			<input class="form-control form-control-sm" type="text" id="bo-exchange-detailid" disabled>
		</div>
			Price paid for item(s): $<span id="bo-exchange-origprice"><%= request.form("price") %></span><br/>
			<label>Price for exchange <% if var_discount > 0 then %>(with discount)<% end if %>:</label><br/>
			<input type="hidden" id="bo-discount-rate" value="<%= var_discount %>">
			<input type="text" id="bo-exchange-price"><br/><br/>
			<span id="bo-exchange-itemname"></span>
			<strong><span id="price-diff-label"></span>$<span id="bo-exchange-price-diff"></span></strong><br/><br/>
			<div id="exchange-agendas"></div>
			<span class="btn btn-sm btn-secondary btn_bo" style="display:none" id="btn-exchange" data-agenda="exchange" data-item="<%= request.form("item") %>" data-detailid="" data-origitem="<%= request.form("detailid") %>" data-exchange_agenda="" data-price="">Set up exchange</span>
		</form>
	</div>
	</div>
<div>
	<div class="float-left" id="exchange-results"></div>
</div>
</div>

</div><!-- row -->
</div><!-- container -->



