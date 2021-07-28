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

<div class="h6 mt-4">Refund item only $<%= request.form("price") %></div>
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



