<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<%

	page_title = "Bodyartforms shopping cart"
	page_description = "Bodyartforms shopping cart"
	page_keywords = "body jewelry, shopping cart, basket"

	' Setting some variables so it doesn't generate string mis match errors and 500 out pages
	session("amount_to_collect") = 0
	session("var_other_items") = 0
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<script type="text/javascript">
// GTM Remove Item from cart
window.dataLayer = window.dataLayer || [];
window.onload=function(){

	var menu = document.querySelector(".action-remove");
menu.addEventListener("click", function(e){
  //  alert("success");
});
		
	var button_removecart = document.querySelector(".action-remove");
	button_removecart.addEventListener("click", function(e){
		console.log("test");
		console.log(e.target.getAttribute("data-productid"));

		var variant = e.target.getAttribute("data-variant");

		window.dataLayer.push({
		event: 'baf.removeFromCart',
		ecommerce: {
			add: {
			products: [{
				id: 'xxxx',
				name: 'xxxx',
				category: 'xxxx',
				variant: variant,
				brand: 'xxxx',
				quantity: 1
				
			}]
			}
		}
		});
	});
} // Run after window finishes loading

</script>	
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<link rel="stylesheet" type="text/css" href="/CSS/slick.css"/>
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="cart/generate_guest_id.asp"-->
<%
check_stock = "yes"

' set page specific variables
session("cart_page") = "yes"

 ' clearing any sessions that could give away free money
Flagged = "" 
var_viewcart_showgifts = "yes"

if request.querystring("remove_save") <> "" then 

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	' cart_save_for_later = 0, initial value
	' cart_save_for_later = 1, it is saved for later
	' cart_save_for_later = 2, it is added back to the cart (to track how many people is checking out with saved for later cart items)
	objCmd.CommandText = "UPDATE tbl_carts SET cart_save_for_later = 2 WHERE cart_id = ? AND " & var_db_field & " = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("cart_id",3,1,10, request.querystring("remove_save")))
	objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10, var_cart_userid))
	objCmd.Execute()
	
end if 

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM TBL_Toggle_Items"
Set rsToggles = objCmd.Execute()

While Not rsToggles.EOF
	If rsToggles("toggle_item") = "toggle_autoclave" Then toggle_autoclave = rsToggles("value")
	If rsToggles("toggle_item") = "toggle_checkout_cards" Then toggle_checkout_cards = rsToggles("value")
	If rsToggles("toggle_item") = "toggle_checkout_paypal" Then toggle_checkout_paypal = rsToggles("value")
	rsToggles.MoveNext
Wend
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="cart/inc_cart_add_item.asp"-->
<!--#include virtual="cart/inc_cart_main.asp"-->
<!--#include virtual="cart/fraud_checks/inc-flagged-orders.asp"-->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<% if request.querystring("addons") = "removed" then %>
<div class="alert alert-success alert-dismissible">
	<h4>Add-on item(s) have been cancelled</h4><button type="button" class="close" data-dismiss="alert" aria-label="Close"><span aria-hidden="true">&times;</span></button>
</div>
<% end if %>
<% if request.querystring("updateditem") = "yes" then %>
<div class="alert alert-success alert-dismissible">
	<h4>Your item has been updated</h4><button type="button" class="close" data-dismiss="alert" aria-label="Close"><span aria-hidden="true">&times;</span></button>
</div>
<% end if %>
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SP_Get_Timer"
Set rsTimer = objCmd.Execute()
%>
<%If NOT rsTimer.EOF Then%>
	<div class="text-center" id="countdown-timer" style="display:none">
		<% countdown_timer = "" & _
		"<div id=""clockdiv"">" & _
		  "<div>" & _
			"<span class=""hours mr-1""></span>: " & _
		  "</div>" & _
		  "<div class=""ml-1"">" & _
			"<span class=""minutes mr-1""></span>: " & _
		  "</div>" & _
		  "<div class=""ml-1"">" & _
			"<span class=""seconds""></span>" & _
		  "</div>" & _
		"</div>"
		%>
		<%= Replace(rsTimer("text_message"), "[Timer]", countdown_timer) %>
	</div>
<% end if %>
<div class="display-5 mb-2">
		Shopping Cart
</div>
<% if session("continue_shopping_link") <> "" then  %>
<a class="btn btn-purple" href="<%= session("continue_shopping_link") %>">Continue shopping</a>
<% end if %>
	
<% ' ------------------------------ BLOCK ACCESS TO PAGE IF FLAGGED ---------------------------- 
IF Flagged = "yes" or session("flag") = "yes" then 
'if 1 <> 1 then
%>
<div class="alert alert-danger"> Access denied -- 
This order can not be processed online. Please contact customer service for assistance.
</div>
<% else %>     

<%
' Show if cart is empty
if cart_status = "empty" Then
%>
<div class="alert alert-primary my-4">There are no items in your shopping cart</div>
<!--#include virtual="cart/inc_stock_display_notice.asp"-->
<%
End If 'End Show if cart is empty

' If customer is NOT registered then clear their cart out of the temp cart DB table


' Show if cart is NOT empty
if cart_status = "not-empty" Then

'====== TRACK THE LAST DATE USER VIEWED THE CART PAGE =================
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE tbl_carts SET cartpage_date_viewed = GETDATE() WHERE " & var_db_field & " = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10, var_cart_userid))
objCmd.Execute()
%> 
<section>
<!--#include virtual="/includes/inc-currency-images.asp" -->
<!--#include virtual="cart/inc_stock_display_notice.asp"-->
<div class="container-fluid mt-5">
	<div class="row">

<div class="col-12 col-lg-8 col-break1600-9 col-break1900-9 pr-lg-5" style="padding-left: .75em;padding-right:0">
<div class="container-fluid p-0" style="margin-left:-.75em;margin-right:-.75em">
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
	<div class="row detailid_<%= rs_getCart.Fields.Item("cart_id").Value %>">
                 <div class="col-auto col-xl-auto mb-3">
				  <% If Instr(rs_getCart.Fields.Item("title").Value, "Digital gift certificate") > 0 Then
					product_link = "gift-certificate.asp"
				  else
					product_link = "productdetails.asp?ProductID=" & rs_getCart.Fields.Item("ProductID").Value
				  end if
				  %>
				  <a href="<%= product_link %>"><div class="position-relative"><img  src="https://bodyartforms-products.bodyartforms.com/<%=(rs_getCart.Fields.Item("picture").Value)%>" alt="Product photo">

					<% ' only display if the item is cheaper than retail 

			if (rs_getCart.Fields.Item("SaleDiscount").Value > 0 AND rs_getCart.Fields.Item("secret_sale").Value = 0) OR (rs_getCart.Fields.Item("secret_sale").Value = 1 AND session("secret_sale") = "yes") then
		%>
			<span class="product-badges badge badge-danger position-absolute rounded-0 p-1">          
			 	 <%= rs_getCart.Fields.Item("SaleDiscount").Value %>% OFF
			  </span>
		<% end if %>
							</div><!-- position-relative -->
				</a>
				 </div><!-- end image -->
				 <div class="col col-lg-9 col-xl-4 small pl-0 mb-3">	
						<%=(rs_getCart.Fields.Item("title").Value)%>

				  <% if rs_getCart.Fields.Item("pair").Value = "yes" then
					var_pair_status = "pair"
						qty_pair_text = "/ pair"
					else
						var_pair_status = "single"
						qty_pair_text = "ea"
					end if 
					%>
					  <div class="font-weight-bold">Sold as a <%= var_pair_status %></div>
					  <% if InStr(rs_getCart.Fields.Item("gauge").Value,"n/a") < 1 then %>
					<div>
				 		<span class="font-weight-bold">Size:</span> <%=(rs_getCart.Fields.Item("gauge").Value)%>
					</div>
					<% end if %>
				  <% if rs_getCart.Fields.Item("ProductDetail1").Value <> "" then %>
					<div>
						<span class="font-weight-bold">Specs:</span> <%=(rs_getCart.Fields.Item("ProductDetail1").Value)%>
					</div>
				  <% end if %>
				  <% if rs_getCart.Fields.Item("length").Value <> "" then %>	  
					<div>
				  		<span class="font-weight-bold">Length:</span>  <%=(rs_getCart.Fields.Item("length").Value)%>
					</div>	
			<% end if %>

			<% if InStr(rs_getCart.Fields.Item("internal").Value,"n/a") < 1 and InStr(rs_getCart.Fields.Item("internal").Value,"null") < 1 and rs_getCart.Fields.Item("internal").Value <> "" then %>	  
				<div>
					<span class="font-weight-bold">Threading:</span> <%= replace(rs_getCart.Fields.Item("internal").Value,","," ")%>
				</div>
			<% end if %>
			<% if rs_getCart.Fields.Item("cart_preorderNotes").Value <> "" then %>	  
				<% if rs_getCart.Fields.Item("ProductID").Value <> 2424 then ' if item is not a gift certificate %>
					<strong><% if rs_getCart("anodID") = 0 then %>Custom color:<% else %>Your specs:<% end if %></strong> <span class="spectext<%= rs_getCart.Fields.Item("cart_id").Value %>"><%= Sanitize(rs_getCart.Fields.Item("cart_preorderNotes").Value) %></span>

					<div class="spec<%= rs_getCart.Fields.Item("cart_id").Value %>" style="display:none">
						<textarea class="form-control form-control-sm my-2 specvalue<%= rs_getCart.Fields.Item("cart_id").Value %>" data-id="<%= rs_getCart.Fields.Item("cart_id").Value %>" rows="10"><%= Sanitize(rs_getCart.Fields.Item("cart_preorderNotes").Value) %></textarea>
						<i class="fa fa-spinner fa-spin fa-2x specspin<%= rs_getCart.Fields.Item("cart_id").Value %>" style="display:none"></i>
					</div>
				
					<div>
						<% if rs_getCart("anodID") = 0 then %>
					<span class="btn btn-sm btn-outline-secondary edit-spec edit<%= rs_getCart.Fields.Item("cart_id").Value %>" data-id="<%= rs_getCart.Fields.Item("cart_id").Value %>">Edit specs</span>
					<% end if %>
					
					<span class="btn btn-sm btn-outline-success updateconfirm<%= rs_getCart.Fields.Item("cart_id").Value %>" style="display:none" data-id="<%= rs_getCart.Fields.Item("cart_id").Value %>"><i class="fa fa-check"></i></span>

					<span class="btn btn-sm btn-outline-success update-spec update<%= rs_getCart.Fields.Item("cart_id").Value %>" style="display:none" data-id="<%= rs_getCart.Fields.Item("cart_id").Value %>">Update specs</span>
					
					<span class="btn btn-sm btn-outline-danger cancel-spec cancel<%= rs_getCart.Fields.Item("cart_id").Value %>" style="display:none" data-id="<%= rs_getCart.Fields.Item("cart_id").Value %>">Cancel</span>
					</div>
					
				<% else ' show gift certificate information 
					certificate_array =split(rs_getCart.Fields.Item("cart_preorderNotes").Value,"{}")				
				%>
				<span class="font-weight-bold">Recipient's name:</span> <%= certificate_array(3) %>
				<span class="font-weight-bold">Recipient's e-mail:</span> <%= certificate_array(0) %>
				<span class="font-weight-bold">Your name:</span> <%= certificate_array(1) %>
				<span class="font-weight-bold">Your message:</span> <%= certificate_array(2) %>
				<%	end if ' detect whether custom order or gift cert %>
			<% end if %>
		
			<% if rs_getCart.Fields.Item("cart_qty").Value <= rs_getCart.Fields.Item("qty").Value then %>
			<% if rs_getCart.Fields.Item("customorder").Value = "yes" then 
			preorder_in_order = "yes"
			%>
					<span class="d-inline-block my-1 bg-info text-white p-2">
						<%= rs_getCart.Fields.Item("preorder_timeframes").Value %> to receive
					</span>	
			<% end if %>
			<%If var_anodizationSubTotal >0 Then %>
				<span class="d-inline-block my-1 bg-info text-white p-2">
					Custom color add-on
				</span>	
			<%End If%>
			<% end if %>
		
      </div><!-- end col / item information -->
			<div class="col-12 col-lg-12 col-xl pt-2 py-xl-0" style="min-width:140px">
<% 
if var_showgifts <> "no" then ' only display on the viewcart page 

if Request.Cookies("ID") <> "" then ' qty select name value changes if logged in or not
		change_id = rs_getCart.Fields.Item("cart_id").Value
	else
		change_id = rs_getCart.Fields.Item("ProductDetailID").Value
	end if

	' don't allow gift certs to change qty
	%>
	<div class="d-inline d-xl-block">	
	<% if var_giftcert = "no" then %>
	Qty: <div class="form-inline d-inline-block">
				<input class="form-control text-center form-control-sm qty_change qty_change_id_<%=(rs_getCart.Fields.Item("ProductDetailID").Value)%>" style="width: 60px" type="tel" maxlength="2" value="<%=(rs_getCart.Fields.Item("cart_qty").Value)%>" name="qty_change_id_<%= change_id %>" id="<%=(rs_getCart.Fields.Item("cart_id").Value)%>" data-detailid="<%=(rs_getCart.Fields.Item("ProductDetailID").Value)%>" data-orig_qty="<%=(rs_getCart.Fields.Item("cart_qty").Value)%>" data-now_item_price="<%= FormatNumber(var_itemPrice, -1, -2, -2, -2) %>" data-retail_item_price="<%= FormatNumber((rs_getCart.Fields.Item("price").Value), -1, -2, -2, -2) %>" data-item_savings="<%= FormatNumber(var_couponLineTotal, -1, -2, -2, -2) %>" data-anodization-subtotal="<%= FormatNumber(var_anodizationSubTotal, -1, -2, -2, -2) %>"  data-anodization-basePrice="<%= FormatNumber(var_anodizationBasePrice, -1, -2, -2, -2) %>">
			</div>
			<div class="btn btn-sm btn-outline-success success_id_<%=(rs_getCart.Fields.Item("ProductDetailID").Value)%>" style="display:none"><i class="fa fa-check"></i></div>
			<input type="hidden" name="orig-qty-<%= rs_getCart.Fields.Item("ProductDetailID").Value %>" value="<%=(rs_getCart.Fields.Item("cart_qty").Value)%>">
		
		<% end if ' if not a gift cert then show qty adjuster 
		%>
		 	  <span class="mr-3 d-inline-block" data-price="<%= FormatNumber(var_itemPrice, -1, -2, -2, -2) %>"> @ <%= exchange_symbol %><%= FormatNumber(var_itemPrice, -1, -2, -2, -2) %><span class="ml-1"><%= qty_pair_text %></span></span>
					<%
					if FormatNumber(var_itemPrice, -1, -2, -2, -2) < FormatNumber(rs_getCart.Fields.Item("price").Value * exchange_rate, -1, -2, -2, -2) then
					%>
					<strike class="mr-1" data-price="<%= FormatNumber(rs_getCart.Fields.Item("price").Value * exchange_rate, -1, -2, -2, -2) %>"><%=exchange_symbol %><%= FormatNumber(rs_getCart.Fields.Item("price").Value * exchange_rate, -1, -2, -2, -2) %></strike>
					<% end if %>					                
<%

else %>
Qty: <%= rs_getCart.Fields.Item("cart_qty").Value %>
<%
end if 'if var_showgifts <> "no" only display on the viewcart page 
%>	
</div><!-- end qty display -->
<%If var_anodizationSubTotal >0 Then %>
	<div class="d-inline d-xl-block">
			<span class="badge badge-info">+ <%= exchange_symbol %><span class="anodization_line_total_<%= rs_getCart.Fields.Item("ProductDetailID").Value %>" data-price="<%= FormatNumber(var_anodizationSubTotal, -1, -2, -2, -2) %>"><%= FormatNumber(var_anodizationSubTotal, -1, -2, -2, -2) %></span> color add-on</span>
	</div>
<%End If%>
<div class="d-inline d-xl-block">
			<span class="font-weight-bold"><%= exchange_symbol %><span class=" line_item_total_<%= rs_getCart.Fields.Item("ProductDetailID").Value %>" data-price="<%= FormatNumber(var_lineTotal, -1, -2, -2, -2) %>"><%= FormatNumber(var_lineTotal, -1, -2, -2, -2) %></span></span>
			<span class="font-weight-bold ml-1">total</span>
		<% if (rs_getCart.Fields.Item("SaleDiscount").Value <> 0 or Session("CouponPercentage") <> "" OR Session("Preferred") = "yes") AND var_giftcert <> "yes" then  ' only display if the item is cheaper than retail 
%>

		<% if rs_getCart.Fields.Item("SaleExempt").Value = 1 AND (Session("Preferred") = "yes" and Session("CouponPercentage") <> "")then %>

						<span class="d-inline-block badge badge-warning p-1 rounded-0">Coupon exempt</span>
						<% 	end if
		 end if%>
			</div><!-- end line total block -->
	</div><!-- end col /  totals and qty box -->
		<div class="col-12 col-lg-12 col-xl-auto pt-3 py-xl-0">
						<span class="btn btn-sm btn-outline-danger mr-2 action-remove" data-detailid="<%= rs_getCart.Fields.Item("cart_id").Value %>" data-specs="<%= change_id %>" data-productid="<%= rs_getCart.Fields.Item("productID").Value %>" data-name="<%= rs_getCart.Fields.Item("title").Value %>" data-category="<%= rs_getCart.Fields.Item("jewelry").Value %>" data-variant="<%= trim(server.htmlencode(rs_getCart.Fields.Item("variant").Value)) %>" data-brand="<%= rs_getCart.Fields.Item("brandname").Value %>" data-qty="<%= rs_getCart.Fields.Item("cart_qty").Value %>"><i class="fa fa-trash-alt"></i></span>
							<span class="btn btn-sm btn-outline-secondary" data-toggle="modal" id="btn-edit-cart-item" data-target="#edit-cart-item" data-productid="<%= rs_getCart.Fields.Item("ProductID").Value %>" data-cartid="<%= rs_getCart.Fields.Item("cart_id").Value %>">Change gauge/color/length</span>
						<% if Request.Cookies("ID") <> "" then ' If customers are registered display links
						%>
						<span class="cart_save_later">
							<button class="btn btn-sm btn-outline-secondary action-save-later" data-detailid="<%= rs_getCart.Fields.Item("cart_id").Value %>" type="button" title="Save for later" >Save for later</button>
						</span>
						<% end if %>
		</div>
</div><!-- end row -->
<hr class="detailid_<%= rs_getCart.Fields.Item("cart_id").Value %>">
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->


	</div><!-- end cart items container -->
	</div><!-- end items column-->
		<div class="col-12 col-lg-4 col-break1600 col-break1900 m-0 p-0">
			<div class="sticky-top" style="z-index:100">
				<div class="card bg-light mb-2">
					
								<div class="card-body text-left py-2">						
									
									<% if Request.Cookies("ID") <> "" then ' if customer is logged in %>
										<% if (rsGetUser.Fields.Item("credits").Value) <> 0 AND session("usecredit") <> "yes" then %>
											  <a class="btn btn-sm btn-outline-secondary my-2 d-block" href="cart.asp?usecredit=yes">Press to use your <%= FormatCurrency(TotalCredits,2) %> store credit</a>
										<% end if 'if customer has a credit to be able to use
									end if %>

									<form id="frm-coupon" action="">
											<div class="input-group input-group-sm my-2">
													<input class="form-control" placeholder="Coupon or certificate code:" name="coupon_code" id="coupon_code" type="text" value="<% if var_display_coupon_code <> "" then  %><%= var_display_coupon_code %><% end if %>" />
												<div class="input-group-append">
														<button id="btn-apply-coupon" class="btn btn-secondary" type="submit">Apply</button>
												</div>
											</div>
									</form>
									<div id="spinner-apply-coupon" style="display:none" class="mb-3"><i class="fa fa-spinner fa-spin fa-lg ml-3"></i> Applying coupon...</div>
									<div id="processing-message"></div>
									<div id="coupon-applied">
										<% if var_display_coupon_code <> "" then %>
											<div class="btn btn-sm btn-block btn-outline-info coupon-shortcut mt-2 font-weight-bold">Press to get <%= var_display_coupon_amount %>% OFF your order!</div>
										<% end if %>
										<div class="row">	
											<div class="col-7">Subtotal</div><div class="col-5">$<span class="cart_subtotal"><%= FormatNumber((var_subtotal), -1, -2, -2, -2) %></span></div>
										</div>	
										<% if Request.Cookies("ID") <> "" then %>
											<% if TotalSpent > 275 AND Session("CouponCode") = "" then %>
												<div class="row">
													<div class="col-7">Your 10% discount</div><div class="col-5">- <span class="cart_prefferred_discount"><%= FormatCurrency(total_preferred_discount, -1, -2, -2, -2) %></span></div>
												</div>
											<% end if ' if preferred customer 
											%>
										<% end if ' if customer is logged in
										%>	
										
										<%
										if discounts_applied = "yes" and Request("coupon_code") <> "" then
										%> 
											<div class="alert alert-success p-1"><%= Valid_type %> APPLIED</div>
										<% end if %>

										<%
										if discounts_applied = "no" and Request("coupon_code") <> ""  then %> 
											<div class="alert alert-danger p-1"><%= Valid_type %> NOT VALID</div>
										<% end if %>
																				
										<% if Session("CouponCode") <> "" then %>
											<div class="row">
												<div class="col-7">Coupon</div><div class="col-5">- $<span class="cart_coupon-amt"><%= FormatNumber(var_couponTotal, -1, -2, -2, -2) %></span></div>
											</div>
										<% end if %>

										<% if Session("GiftCertAmount") <> 0 then %>
											<div id="row_gift_cert">
												<div class="row">
													<div class="col-7">Gift certificate</div><div class="col-5">- <span id="cart_gift-cert"><%= FormatCurrency(Session("GiftCertAmount"), -1, -2, -2, -2) %></span></div>
												</div>
											</div>
										<% ' if there is a gift certificate found
										end if 
										%>										
									</div>									
									
									<div id="row_use_now_credits">
										<div class="row">
											<div class="col-7">Order credits</div><div class="col-5">- <span id="use_now_amount"><%= FormatCurrency(credit_now,2) %></span></div>
										</div>
									</div>
									<% if session("usecredit") = "yes" then %>
									<div id="row_store_credit">
										<div class="row">
											<div class="col-7">Store credit</div><div class="col-5">- <span id="store_credit_amt"><%= FormatCurrency(session("storeCredit_used"),2) %></span><span title="Remove store credit" id="remove-credit" class="text-danger ml-3 pointer" data-type="store-credit"><i class="fa fa-trash-alt"></i></span>
											</div>
										</div>	
									</div>
									<% end if 'if customer has a credit to be able to use %>	
								
								
									<% 
									if Request.ServerVariables("URL") = "/cart.asp" or Request.ServerVariables("URL") = "/cart2.asp" then
											est_shipping = "Est shipping"
										else
											est_shipping = "Shipping"
										end if
									%>
									
									<div class="row">
										   <div class="col-7"><%= est_shipping %></div><div class="col-5 cart_shipping"><%= var_shipping_cost_friendly %></div>
									</div>
									<div class="row">
										<div class="col-7">Tax</div><div class="col-5 small">Calculated on next screen</div>
									</div>
									
									<% ' do not show free shipping notice if order is heavy 
									if session("weight") <= 32 and strcountryName = "US" and var_other_items = 1 and request.cookies("OrderAddonsActive") = "" then
									%>
										<div class="cart_shipping_amountLeft text-center text-success p-1 mt-1" <% if var_shipping_AmountNeeded <= 0 then %>style="display:none"<% end if %>>
												<i class="fa fa-shipping-fast fa-lg mr-2"></i>
												<span class="font-weight-bold">Only <span class="shipping_amount_left"><%= FormatCurrency(var_shipping_AmountNeeded, -1, -2, -2, -2) %></span> away from <%= var_shipping_goal %> SHIPPING</span>
										</div>
									<% end if ' free shipping notice only showing if order is not heavy
									%>	
									
									<%If var_other_items = 1 then 
										If request.cookies("OrderAddonsActive") = "" then%>
											<div id="stepper"></div>	
										<%End If		
									End If%>

								</div><!-- end card body -->
								
								<div class="card-footer">
												
									<!--#include virtual="cart/inc_cart_grandtotal.asp"-->
									<h4>TOTAL <% if strcountryName <> "US" then %> (USD)<% end if %>$<span class="cart_grand-total"><%= FormatNumber(var_grandtotal, -1, -2, -2, -2) %></span></h4>
										<div class="row_convert_total" style="display:none">
												<div class="alert alert-success p-2">
													<div><h6><img class="mr-2" style="width:20px;height:20px" src="/images/icons/<%= currency_img %>">ESTIMATE <span class="exchange-price"><span class="currency-type"></span> <span class="convert-total convert-price" data-price=""></span></span></h6></div>
														<span class="exchange-price"><span class="currency-type bold"></span> <span class="convert-total convert-price bold" data-price=""></span> is a close estimate</span>. The total billed will be for <span class="bold">$<span class="cart_grand-total"><%= FormatNumber(var_grandtotal, -1, -2, -2, -2) %></span> in US Dollars</span> and your bank will convert to the most current exchange rate.
												</div>
										</div>
										<% if var_anodization_added = 1 then 
										%>
												<div class="bg-warning p-2 rounded my-2 text-center">
													<i class="fa fa-alert fa-2x"></i><div class="small font-weight-bold">Custom color added<br>Order will be delayed by up to 1 week</div>
												</div>	
										<% end if %>
									<% If toggle_checkout_cards = true Then %>
											<div class="checkout_now" style="display:none">
												<a id="btn-checkout" class="btn btn-block btn-primary mb-2 checkout_button" href="checkout.asp?type=card" ><h6>CHECKOUT WITH <span class="payment-options">CREDIT CARD
													<br/>
													<span style="font-size:2em">
													<i class="fa fa-cc-visa"></i>
													<i class="fa fa-cc-mastercard"></i>
													<i class="fa fa-cc-amex"></i>
													<i class="fa fa-cc-discover"></i></span>
												</span>
												</h6>
												</a>
											</div>
									<% else %>
											<div class="alert alert-danger">We're sorry, but our <b>credit card</b> checkout is temporarily unavailable. As soon as our payment processor comes back online, we will accept orders again. Please check back later.</div>
									<% end if %>										

									<% If toggle_checkout_paypal = true Then %>
										<div class="checkout_paypal mb-2"  style="display:none">
											<a id="btn-paypal" class="btn btn-block btn-warning checkout_button" href="checkout.asp?type=paypal">
												<img style="height:30px" src="/images/paypal.png" />
											</a>
										</div>
									<% else %>
										<div class="alert alert-danger">We're sorry, but our <b>PayPal</b> checkout is temporarily unavailable. As soon as PayPal comes back online, we will accept orders again. Please check back later.</div>
									<% end if %>
									<div id="pay-api-processing-message" style="display:none"></div>
									<div id="pay-api-error-message" class="alert alert-danger" style="display:none"></div>	
									<apple-pay-button id="btn-applepay" buttonstyle="<%= var_applepay_color %>" type="check-out" locale="en" class="mb-2"></apple-pay-button>
									<div id="btn-googlepay" class="mb-3 checkout_button" style="width: 100%; height: 45px; display: none;"></div>
									
									
									<%
									' === only show afterpay option to USA customers
									if request.cookies("currency") = "" OR request.cookies("currency") = "USD" then
										afterpay_display = ""
									else
										afterpay_display = "display:none"
									end if
									%>
									<div id="REMOVE-GO-LIVE" style="display:none">
										<div class="afterpay_option" style="<%= afterpay_display %>">
											<a class="btn btn-block btn-outline-secondary pb-1  mt-3 " style="display:none" id="btn-afterpay-checkout" href="checkout.asp?type=afterpay"><span class="afterpay-widget"></span></a>
											<div class="mt-3" style="display:none" id="afterpay-displayonly"><span class="afterpay-widget-nonactive"></span></div>
										</div>
									</div>
						</div><!-- end card footer for totals -->
					  </div><!-- end card for totals -->
<% 
'===== CHECK STOCK ON PRODUCTS BEING OFFERED AS ADDONS AT CHECKOUT
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT qty, title, picture, ProductDetailID, price, jewelry.ProductID FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID WHERE qty >= 10 and (jewelry.ProductID = 28568 OR ProductDetailID = 166942)"
set rsGetAddOns = objCmd.Execute()

if NOT rsGetAddOns.eof then
%>
						<div class="text-center mb-2 mt-3"><h5>Need aftercare salve?</h5></div>
						<div class="row mb-3 mx-0 p-0">
					<%					
					while NOT rsGetAddOns.eof
					%>
					<div class="col-6">
						<a href="/productdetails.asp?ProductID=<%= rsGetAddOns("ProductID") %>"><img class="img-fluid text-left pull-left rounded-circle mr-2 mb-2" src="https://bodyartforms-products.bodyartforms.com/<%= rsGetAddOns("picture") %>"></a>
						<button class="btn btn-sm btn-purple add-cart-addon" data-detailid="<%= rsGetAddOns("ProductDetailID") %>" id="btn_<%= rsGetAddOns("ProductDetailID") %>">Add to cart</button>
						<br>
						<%= rsGetAddOns("title") %><br>
						<%= FormatCurrency(rsGetAddOns("price"),2) %>
						
						
					</div>
					<%
					rsGetAddOns.movenext()
					Wend
					%>
						</div><!-- addons row -->
<% end if ' if NOT rsGetAddOns.eof
%>
									<% ' Display if ANY autoclavable items are found on the order 
									if var_autoclavable = 1 and var_sterilization_added = 0 and toggle_autoclave = true then
									%>
										<div class="alert alert-info p-2 mt-4">
											<div class="clearfix"><img class="float-left mr-2" src="https://bodyartforms-products.bodyartforms.com/autoclave-2-1464.jpg" style="width: 120px; height:auto">
											<h6>Sterilize the items below for only $4.95</h6>
											<span class="btn btn-sm btn-outline-info btn-add-autoclave">Add service to cart</span>
											</div>
												
												<ul class="my-2">
												<% if str_autoclave_items <> "" then
												%>
													<%= str_autoclave_items %>
												<% 
													end if ' str_autoclave_items
												%>
												</ul>
												<div class="small">
													Adding on autoclave sterilization service will only delay your order by 1 business day (Express orders will not be delayed). 
												</div>
												
									
									</div> 
									<% end if %>
								
				</div><!-- sticky top -->
		</div><!-- totals column -->
</div><!-- entire cart row -->
</div><!-- entire cart container -->


	</section>
		<!-- Update cart item Modal -->
        <div class="modal fade" id="edit-cart-item" tabindex="-1" role="dialog"                 aria-hidden="true">
                <div class="modal-dialog" role="document">
                        <div class="modal-content">
                                <div class="modal-header">
                                    <h5 class="modal-title">Change item gauge, length, or style/color</h5>
                                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                                        <span aria-hidden="true">&times;</span>
                                    </button>
                                </div>
								<div class="modal-body">
									<div id="update-item-display-product"></div>
									<span class="clearfix"></span>
									<form id="form-edit-cart-item"></form>
								</div>  
								<div class="modal-footer">
									<div class="d-inline-block text-right w-50">
										<button type="button" class="btn btn-sm btn-purple"  id="btn-update-detail">Update item</button>
									</div>
								</div>  
                        </div>
                </div>
        </div>	
			
			<!-- PRE_ORDER ITEM WARNING -->
			<div class="modal fade" id="custom-order-warning-modal" tabindex="-1" role="dialog" aria-hidden="true">
				<div class="modal-dialog" role="document">
						<div class="modal-content">
								<div class="modal-header">
									<h5 class="modal-title">Your order contains custom made items</h5>
									<button type="button" class="close" data-dismiss="modal" aria-label="Close">
										<span aria-hidden="true">&times;</span>
									</button>
								</div>
								<div id="custom-order-items" class="modal-body">

								</div>
								<button id="btn-proceed-to-checkout" class="btn btn-lg btn-primary btn-block proceed_checkout w-75 mx-auto mb-3" type="button">PROCEED TO CHECKOUT</button>
								<!--
								<div class="modal-footer">
									<div class="d-inline-block text-right w-50">
										<button type="button" class="btn btn-secondary close-bo" data-dismiss="modal">Cancel</button>
									</div>
								</div>  
								-->
						</div>
				</div>
			</div>	

			<!-- FREE ITEMS MODAL -->
			<div class="modal fade" id="free-items-modal" tabindex="-1" role="dialog" aria-hidden="true">
				<div class="modal-dialog" role="document">
						<div class="modal-content">
								<div class="modal-header">
									<h5 class="modal-title">SELECT FREE ITEMS</h5>
									<button type="button" class="close" data-dismiss="modal" aria-label="Close">
										<!--<span aria-hidden="true">&times;</span>-->
									</button>
								</div>
								<div id="loading-message" class="modal-body"></div>
								<div id="free-items" class="modal-body">
									
								</div>
								<div class="modal-footer">
									<div class="d-inline-block text-right w-50">
										<button type="button" class="btn btn-secondary close-bo" data-dismiss="modal">Done</button>
									</div>
								</div>  
						</div>
				</div>
			</div>	
			
			<!-- FREE ITEMS PAGE MODAL -->
			<div class="modal fade" id="free-items-page-modal" tabindex="-1" role="dialog" aria-hidden="true">
				<div class="modal-dialog modal-xl" role="document">
						<div class="modal-content">
								<div class="modal-header">
									<h5 class="modal-title">Free Items</h5>
									<button type="button" class="close" data-dismiss="modal" aria-label="Close">
										<span aria-hidden="true">&times;</span>
									</button>
								</div>
								<div class="modal-body">
									<!--#include virtual="/free-items-modal.asp" -->
								</div>
						</div>
				</div>
			</div>	
			
			<!-- Free items page / Item variations modal -->
			<div class="modal fade" id="free-items-page-variations-modal" role="dialog" aria-labelledby="headVariation" aria-hidden="true">
				<div class="modal-dialog modal-sm" role="document">
				  <div class="modal-content">
					<div class="modal-body">
							<div id="load-variants"><i class="fa fa-spinner fa-2x fa-spin"></i></div>
					</div>
					<div class="modal-footer">
					  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
					</div>
				  </div>
				</div>
			  </div>
			<!-- end variations modal -->
	
			
	<%
	End If 'End Of cart show if not empty
	%>
	<% if CustID_Cookie <> 0 then %>
	<div class="mt-5" id="saved-items"></div>
	<% end if %>


	<% end if 'block access to page if user is flagged %>
<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript" src="/js-pages/toggle_required_billing.js"></script>
<script type="text/javascript" src="/js-pages/currency-exchange.min.js?v=050619"></script>
<% if (session("exchange-rate") = "" OR session("exchange-currency") <> request.cookies("currency")) AND request.cookies("currency") <> "" AND request.cookies("currency") <> "USD" then %>
<script>
		// Get currency conversions on page load
		updateCurrency();
</script>
<% end if %>
 
<!-- Global variables for Apple & Google Pay -->
<script>
	var tax = 0.0;
	var shippingCost = 0.0;
	var subTotal = 0.0;
	var totalAmount = 0.0;
	var totalDiscount = 0.0; // Gets updated in calcAllTotals()
	var selectedShippingId = 0;
	var selectedShippingCompany = "";
	var preOrderItem = "";
	var checkoutMethod = "";
</script>

<!-- Google Pay Javascript -->
<script src="/js/google-pay-v2api.js?ver=1"></script>
<script async src="https://pay.google.com/gp/p/js/pay.js" onload="onGooglePayLoaded()"></script>

<!-- Apple Pay Javascript -->
<script src="https://applepay.cdn-apple.com/jsapi/v1/apple-pay-sdk.js"></script>
<script src="/js/apple-pay-api.js?ver=3"></script>
<script type="text/javascript" src="/js/slick-customized.min.js"></script>

<!-- !!!!!!!!!!!!!!!!!!!!!  BE SURE TO ALSO UPDATE THE CART JS FILE ON CHECKOUT PAGE !!!!!!!!!!!!!!!!!!!!! -->
<script type="text/javascript" src="/js-pages/cart.min.js?v=03032025"></script>
<script type="text/javascript" src="/js-pages/cart_update_totals.min.js?v=111723"></script>
<!-- ^^^^^^  BE SURE TO ALSO UPDATE THE CART JS FILE ON CHECKOUT PAGE ^^^^^^ -->
<script type="text/javascript">
	calcAllTotals();
</script>
<script type="text/javascript">
// START load variations into modal window pop up
	$(document).on("click", ".view-variations", function(event){
		var productid = $(this).attr("data-id");	
		$('#load-variants').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');
		$('#load-variants').load("products/ajax-freeplanner-getvariations.asp", {productid: productid}, function() {
		});	
	});		
</script>	
<!-- Shipping Countdown Timer -->
<%If Not rsTimer.EOF Then%>
	<script>
		var total_seconds = <%=rsTimer("remaining_seconds")%>;
	</script>
	<script src="/js/countdown-timer.js"></script>
<%End If%>
<%
Set rsToggles = Nothing
Set rsTimer = Nothing
%>
