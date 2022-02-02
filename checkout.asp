<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<%
	page_title = "Bodyartforms checkout"
	page_description = "Bodyartforms checkout"
	page_keywords = ""
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<% 
' set page specific variables
session("cart_page") = "no"
check_stock = "yes"
%>
<!--#include file="functions/encrypt.asp"-->
<!--#include file="Connections/authnet.asp"-->
<!--#include virtual="cart/inc_cart_main.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<%
' KLAVIYO SCRIPT BEGIN 
product_names = product_names & """" & rs_getCart("title") & ""","
categories = categories & """" & Trim(rs_getCart("jewelry")) & ""","
products = products & "{" & _
         "'ProductID': '" & rs_getCart("ProductID") & "'," & _
         "'SKU': '" & rs_getCart("cart_detailID") & "'," & _
         "'ProductName': '" & rs_getCart("title") & "'," & _
         "'Quantity': " & rs_getCart("cart_qty") & "," & _
         "'ItemPrice': " & rs_getCart("price") & "," & _
         "'RowTotal': " & rs_getCart("cart_qty") * rs_getCart("price") & "," & _
         "'ProductURL': 'https://bodyartforms.com/productdetails.asp?productid=" & rs_getCart("ProductID") & "'," & _
         "'ImageURL': 'https://bodyartforms-products.bodyartforms.com/" & rs_getCart("picture") & "'," & _
         "'ProductCategories': ['" & Trim(rs_getCart("jewelry")) & "']" & _
       "},"
%>
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<%
'Remove last coma
If products <> "" Then products = Mid(products, 1, LEN(products)-1)
If product_names <> "" Then product_names = Mid(product_names, 1, LEN(product_names)-1)
If categories <> "" Then categories = Mid(categories, 1, LEN(categories)-1)
%>
<script>
	//Klaviyo Started Checkout push
	_learnq.push(["track", "Started Checkout", {
	 "$event_id": "<%= var_cart_userid %>_<%= Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) %>",
     "$value": <%=var_subtotal%>,
     "ItemNames": [<%=product_names%>],
     "CheckoutURL": "https://bodyartforms.com/checkout.asp",
     "Categories": [<%=categories%>],
     "Items": [<%=products%>]
   }]);
</script>
<!-- KLAVIYO SCRIPT END -->
<!--#include virtual="cart/inc_cart_grandtotal.asp"-->
<%
'response.write "<br/>subtotal " & var_subtotal
'response.write "<br/>var_salesTax " & var_salesTax
%>
<!--#include virtual="cart/fraud_checks/inc-flagged-orders.asp"-->

<% 
' Set variable to have free gifts show values without showing dropdowns
var_showgifts = "no"

' Set country session based on IP geolocation targeting
if request.form("shipping-country") = "" then
	if strcountryName = "US" then
		session("shipping-country") = "USA"
		session("shipping-display") = "yes"
	end if
	if strcountryName = "CA" then
		session("shipping-country") = "Canada"
		session("shipping-display") = "yes"
	end if
end if
'----------------

Set rsGetCountrySelect = Server.CreateObject("ADODB.Recordset")
rsGetCountrySelect.ActiveConnection = DataConn
rsGetCountrySelect.Source = "SELECT * FROM dbo.TBL_Countries WHERE Display = 1 ORDER BY Country ASC "
rsGetCountrySelect.CursorLocation = 3 'adUseClient
rsGetCountrySelect.LockType = 1 'Read-only records
rsGetCountrySelect.Open()

if CustID_Cookie <> 0 then 
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandType = 4
	objCmd.CommandText = "SP_inc_GetCustomer_ByCookie"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
	Set rsGetUser = objCmd.Execute()

	'Assign user information to variables
	If Not rsGetUser.EOF Or Not rsGetUser.BOF Then
		' Set session variable to modify shipping/billing information in account
		session("cim_accountNumber") = rsGetUser.Fields.Item("cim_custid").Value
	End if
end if ' if user is logged in

'Hide/show classes if registered or not registered
if CustID_Cookie = 0 then
	hide_non_registered = "style=""display:none"""
else
	hide_registered = "style=""display:none"""
end if


' no display states/provinces based on country type from viewcart page setting
if session("shipping-country") = "USA" then
	hide_canada = "style=""display:none"""
	hide_province = "style=""display:none"""
	hide_inter_zip = "style=""display:none""" 
end if
if session("shipping-country") = "Canada" then
	hide_state = "style=""display:none"""
	hide_province = "style=""display:none"""
end if
if session("shipping-country") <> "Canada" AND session("shipping-country") <> "USA" then
	hide_state = "style=""display:none"""
	hide_canada = "style=""display:none"""
end if


	if request.cookies("gaugecard") <> "no" then
		free_card = " OR ProductDetailID = 5461 "
	end if
	if request.cookies("oringsid") <> "" then
		free_orings = " OR ProductDetailID = ? "
	end if
	if request.cookies("stickerid") <> "" then
		free_sticker = " OR ProductDetailID = ? "
	end if
	if request.cookies("freegift1id") <> "" then
		free_gift1 = " OR ProductDetailID = ? "
	end if
	if request.cookies("freegift2id") <> "" then
		free_gift2 = " OR ProductDetailID = ? "
	end if
	if request.cookies("freegift3id") <> "" then
		free_gift3 = " OR ProductDetailID = ? "
	end if
	if request.cookies("freegift4id") <> "" then
		free_gift4 = " OR ProductDetailID = ? "
	end if
	if request.cookies("freegift5id") <> "" then
		free_gift5 = " OR ProductDetailID = ? "
	end if
	'response.write " free card ---" & free_card & "---<br/>"
	'response.write " free o-rings ---" & free_orings & "---<br/>"
	'response.write " free sticker ---" & free_sticker & "---<br/>"
	'response.write " free free_gift1 ---" & free_gift1 & "---<br/>"
	'response.write " free free_gift2 ---" & free_gift2 & "---<br/>"
	'response.write " free free_gift3 ---" & free_gift3 & "---<br/>"
	'response.write " free_gift4 ---" & free_gift4 & "---<br/>"
	'response.write " free_gift5 ---" & free_gift5 & "---<br/>"
	
	free_result = Mid(free_card & free_orings & free_sticker & free_gift1 & free_gift2 & free_gift3 & free_gift4 & free_gift5, 5)

	'response.write "post trim ---" & free_result & "---"
	if free_result <> "" then
		free_result = " WHERE " & free_result
	else
		free_result = ""
	end if

	if free_result <> "" then
	

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT jewelry.title, jewelry.picture, ProductDetails.ProductDetail1, ProductDetails.qty, ProductDetails.free, jewelry.ProductID, ProductDetails.ProductDetailID, ProductDetails.Free_QTY, ProductDetails.weight, jewelry.picture, ProductDetails.price, ProductDetails.active,  ProductDetails.Gauge, ProductDetails.Length, ProductDetails.detail_code FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID " & free_result & " ORDER BY ProductDetailID ASC"

		if request.cookies("oringsid") <> "" then
			objCmd.Parameters.Append(objCmd.CreateParameter("orings",3,1,10,request.cookies("oringsid")))
		end if
		if request.cookies("stickerid") <> "" then
			objCmd.Parameters.Append(objCmd.CreateParameter("sticker",3,1,10,request.cookies("stickerid")))
		end if
		if request.cookies("freegift1id") <> "" then
			objCmd.Parameters.Append(objCmd.CreateParameter("gift1",3,1,10,request.cookies("freegift1id")))
		end if
		if request.cookies("freegift2id") <> "" then
			objCmd.Parameters.Append(objCmd.CreateParameter("gift2",3,1,10,request.cookies("freegift2id")))
		end if
		if request.cookies("freegift3id") <> "" then
			objCmd.Parameters.Append(objCmd.CreateParameter("gift3",3,1,10,request.cookies("freegift3id")))
		end if
		if request.cookies("freegift4id") <> "" then
			objCmd.Parameters.Append(objCmd.CreateParameter("gift4",3,1,10,request.cookies("freegift4id")))
		end if
		if request.cookies("freegift5id") <> "" then
			objCmd.Parameters.Append(objCmd.CreateParameter("gift5",3,1,10,request.cookies("freegift5id")))
		end if

	Set rsFreeGifts = objCmd.Execute()

While not rsFreeGifts.eof
	var_free_item = ""
	detailid = int(rsFreeGifts.Fields.Item("ProductDetailID").Value & "0")

	if rsFreeGifts.Fields.Item("ProductDetailID").Value = 5461 then
		var_free_item = "<span class=""mr-2"">FREE:</span>Gauge card<br/>"
	end if
	if detailid = int(request.cookies("oringsid") & "0") then
		var_free_item = "FREE: <span class=""ml-1 mr-2"">Qty 4</span> " & rsFreeGifts.Fields.Item("gauge").Value & " " & rsFreeGifts.Fields.Item("length").Value & " " &   rsFreeGifts.Fields.Item("ProductDetail1").Value &  " " & rsFreeGifts.Fields.Item("title").Value & "<br/>"
	end if
	if detailid = int(request.cookies("stickerid") & "0") then
		var_free_item = "<span class=""mr-2"">FREE:</span>" & rsFreeGifts.Fields.Item("ProductDetail1").Value &  " " & rsFreeGifts.Fields.Item("title").Value & "<br/>"
	end if

	var_free_items = var_free_items & var_free_item

	if detailid = int(request.cookies("freegift1id") & "0") then
		var_free_item = "FREE: <span class=""ml-1 mr-2"">Qty " & rsFreeGifts.Fields.Item("Free_QTY").Value & "</span> " & rsFreeGifts.Fields.Item("gauge").Value & " " & rsFreeGifts.Fields.Item("length").Value & " " &   rsFreeGifts.Fields.Item("ProductDetail1").Value &  " " & rsFreeGifts.Fields.Item("title").Value & "<br/>"
		var_free_items = var_free_items & var_free_item
	end if
	if detailid = int(request.cookies("freegift2id") & "0") then
		var_free_item = "FREE: <span class=""ml-1 mr-2"">Qty " & rsFreeGifts.Fields.Item("Free_QTY").Value & "</span> " & rsFreeGifts.Fields.Item("gauge").Value & " " & rsFreeGifts.Fields.Item("length").Value & " " &   rsFreeGifts.Fields.Item("ProductDetail1").Value &  " " & rsFreeGifts.Fields.Item("title").Value & "<br/>"
		var_free_items = var_free_items & var_free_item
	end if
	if detailid = int(request.cookies("freegift3id") & "0") then
		var_free_item = "FREE: <span class=""ml-1 mr-2"">Qty " & rsFreeGifts.Fields.Item("Free_QTY").Value & "</span> " & rsFreeGifts.Fields.Item("gauge").Value & " " & rsFreeGifts.Fields.Item("length").Value & " " &   rsFreeGifts.Fields.Item("ProductDetail1").Value &  " " & rsFreeGifts.Fields.Item("title").Value & "<br/>"
		var_free_items = var_free_items & var_free_item
	end if
	if detailid = int(request.cookies("freegift4id") & "0") then
		var_free_item = "FREE: <span class=""ml-1 mr-2"">Qty " & rsFreeGifts.Fields.Item("Free_QTY").Value & "</span> " & rsFreeGifts.Fields.Item("gauge").Value & " " & rsFreeGifts.Fields.Item("length").Value & " " &   rsFreeGifts.Fields.Item("ProductDetail1").Value &  " " & rsFreeGifts.Fields.Item("title").Value & "<br/>"
		var_free_items = var_free_items & var_free_item
	end if
	if detailid = int(request.cookies("freegift5id") & "0") then
		var_free_item = "FREE: <span class=""ml-1 mr-2"">Qty " & rsFreeGifts.Fields.Item("Free_QTY").Value & "</span> " & rsFreeGifts.Fields.Item("gauge").Value & " " & rsFreeGifts.Fields.Item("length").Value & " " &   rsFreeGifts.Fields.Item("ProductDetail1").Value &  " " & rsFreeGifts.Fields.Item("title").Value & "<br/>"
		var_free_items = var_free_items & var_free_item
	end if
	
rsFreeGifts.Movenext()
wend

end if ' free_result <> ""

if request.cookies("OrderAddonsActive") <> "" then
	session("invoiceid") = request.cookies("OrderAddonsActive")
	hide_section_addons = "style=""display:none"""
	show_section_addons = "d-block"
else
	show_section_addons = "d-none"
end if 

'Check Toogle Items
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
<!--#include virtual="/includes/inc-currency-images.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<div class="display-5 mb-3">
		Checkout
</div>


<form class="needs-validation" id="checkout_form" novalidate>
		<div class="container-fluid">
				<div class="row">
			
			<div class="col-12 col-lg-8 col-break1600-9 col-break1900-9 pr-lg-5" style="padding-left: .75em;padding-right:0">
			<div class="container-fluid p-0" style="margin-left:-.75em;margin-right:-.75em">
<%
' Show if cart is empty
if cart_status = "empty" Then
%>
	 <div class="alert alert-primary my-3">There are no items in your shopping cart</div>
    <%
End If 'End Show if cart is empty

' If customer is NOT registered then clear their cart out of the temp cart DB table


' Show if cart is NOT empty
if cart_status = "not-empty" Then

'====== TRACK THE LAST DATE USER VIEWED THE CART PAGE =================
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE tbl_carts SET checkoutpage_date_viewed = GETDATE() WHERE " & var_db_field & " = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10, var_cart_userid))
objCmd.Execute()
%> 

 	<% ' ------------------------------ BLOCK ACCESS TO PAGE IF FLAGGED ---------------------------- 
	IF Flagged = "yes" or session("flag") = "yes" then 
	'if 1 <> 1 then 
	%>
	<div class="alert alert-danger my-3"> Too many checkout attempts<br>
	This order can not be processed online. Please contact customer service for assistance.
	</div>
	<% else ' if order or account is not flagged
	%>  

 <div id="msg-location-replace"></div>
 <!--#include virtual="cart/inc_stock_display_notice.asp"-->
 <section> 
<div class="card mb-5" id="shipping-card">
	<div class="card-header">
		<h5>Shipping address</h5>
	</div>
	<div class="card-body">
			<div class="<%= show_section_addons %>">
				Items will be shipped to the address given when order was placed
			</div>
		<div <%= hide_section_addons %>>
			<button class="btn btn-sm btn-info add-new-shipping-button" style="display:none" type="button">Add a new  shipping address</button>
		<button class="btn btn-sm btn-outline-danger mb-4" style="display:none" type="button" id="cancel-shipping-add">Cancel Add / Update</button>
	<div style="display:none" id="cim_shipping_addresses">
		<!--#include virtual="/checkout/cim_getshipping_addresses.asp"-->
	</div>
<div class="shipping-address-form">
<% if CustID_Cookie = 0 then
'and (var_no_ship_addresses = "false" or var_no_ship_addresses = "")
 %>
<br/>
<div class="form-group position-relative" <%= hide_registered %>>
<label for="e-mail">E-mail address <span class="text-danger">*</span></label>
<input class="form-control" required name="e-mail" id="e-mail" type="email" autocomplete="shipping email" data-friendly-error="E-mail address is required" />
<div class="invalid-feedback">
		E-mail address is required
</div>
<div class="invalid-feedback feedback-icon">
	<i class="fa fa-times"></i>
</div>
<div class="valid-feedback feedback-icon">
	<i class="fa fa-check"></i>
</div>
</div>
<% else %>
<input id="e-mail" type="hidden" value="<%= rsGetUser.Fields.Item("email").Value %>" />
<% end if %>

<div class="form-group position-relative">
	<label for="shipping-first-checkout">First name <span class="text-danger">*</span></label>
	<input class="form-control" required name="shipping-first" id="shipping-first-checkout" type="text" autocomplete="shipping given-name" data-friendly-error="Shipping address: First name is required" />
	<div class="invalid-feedback">
		First name is required
	</div>
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
	<div class="invalid-feedback feedback-icon">
		<i class="fa fa-times"></i>
	</div>
</div>
<div class="form-group position-relative">
<label for="shipping-last-checkout">Last name <span class="text-danger">*</span></label>
<input class="form-control" required name="shipping-last" id="shipping-last-checkout" type="text" autocomplete="shipping family-name" data-friendly-error="Shipping address: Last name is required"  />
<div class="invalid-feedback">
	Last name is required
</div>
<div class="invalid-feedback feedback-icon">
	<i class="fa fa-times"></i>
</div>
<div class="valid-feedback feedback-icon">
	<i class="fa fa-check"></i>
</div>
</div>
<div class="form-group position-relative">
	<label for="shipping-company">Company</label>
	<input class="form-control" name="shipping-company" id="shipping-company" type="text" autocomplete="shipping organization" />
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
</div>
<div class="form-group position-relative">
	<label for="shipping-address">Address (Line 1)<span class="text-danger">*</span></label>
	<input class="form-control" required name="shipping-address" id="shipping-address" type="text" autocomplete="shipping address-line1" data-friendly-error="Shipping address is required" />
	<div class="invalid-feedback">
		Address is required
	</div>
	<div class="invalid-feedback feedback-icon">
		<i class="fa fa-times"></i>
	</div>
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
</div>
<div class="form-group position-relative">
	<label for="shipping-address2">Apt #, Dorm, Suite</label>
	<input class="form-control" name="shipping-address2" id="shipping-address2" type="text" autocomplete="shipping address-line2" />
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
</div>
<div class="form-group position-relative">
	<label for="shipping-city">City <span class="text-danger">*</span></label>
	<input class="form-control" required name="shipping-city" id="shipping-city" type="text" autocomplete="shipping address-level2" data-friendly-error="Shipping city is required" />
	<div class="invalid-feedback">
		City is required
	</div>
	<div class="invalid-feedback feedback-icon">
		<i class="fa fa-times"></i>
	</div>
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
</div>
<div class="form-group position-relative">                     
	<label for="shipping-country">Country <span class="text-danger">*</span></label>
	<select class="form-control" required name="shipping-country" id="shipping-country" autocomplete="shipping country"  data-friendly-error="Shipping country is required" >
	<option value="USA" selected>USA</option>
	<% 
	While NOT rsGetCountrySelect.EOF 
	%>
			<option value="<%=(rsGetCountrySelect.Fields.Item("Country").Value)%>"><%=(rsGetCountrySelect.Fields.Item("Country").Value)%></option>
	<% 
	rsGetCountrySelect.MoveNext()
	Wend
	rsGetCountrySelect.Requery
	%>
		</select>
		<div class="invalid-feedback">
			Country is required
		</div>
		<div class="invalid-feedback feedback-icon">
			<i class="fa fa-times"></i>
		</div>
		<div class="valid-feedback feedback-icon">
			<i class="fa fa-check"></i>
		</div>
</div>
<div class="form-group shipping-state position-relative" <%= hide_state %>>
	<label for="shipping-state">State (USA)<span class="text-danger">*</span></label>
	<select class="form-control" required name="shipping-state" id="shipping-state" autocomplete="shipping address-level1"  data-friendly-error="Shipping state is required">
	<!--#include file="includes/inc_states_select.asp"-->
      </select>
	  <div class="invalid-feedback">
		State is required
	</div>
	<div class="invalid-feedback feedback-icon">
		<i class="fa fa-times"></i>
	</div>
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
</div>
<div class="form-group shipping-province-canada position-relative" <%= hide_canada %>>
	<label for="shipping-province-canada">Province <span class="text-danger">*</span></label>
	<select class="form-control" name="shipping-province-canada" id="shipping-province-canada" autocomplete="shipping address-level2" data-friendly-error="Shipping province is required">
	<!--#include virtual="/includes/inc_province_canada_select.asp"-->
	  </select>
	  <div class="invalid-feedback">
		Province is required
	</div>
	<div class="invalid-feedback feedback-icon">
		<i class="fa fa-times"></i>
	</div>
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
</div>
<div class="form-group shipping-province position-relative" <%= hide_province %>>
	<label for="shipping-province">Province / State</label>
	<input class="form-control" name="shipping-province" id="shipping-province" type="text" autocomplete="shipping address-level2" data-friendly-error="Shipping province/state is required" />
	<div class="invalid-feedback">
		Province / State is required
	</div>
	<div class="invalid-feedback feedback-icon">
		<i class="fa fa-times"></i>
	</div>
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
</div>

<div class="form-group position-relative">
	<label for="shipping-zip"><span class="hide_usa_zip <%= hide_state %>">Zip code</span><span class="hide_inter_zip <%= hide_inter_zip %>">Postal code</span> <span class="text-danger">*</span></label>
	<input class="form-control" required name="shipping-zip" id="shipping-zip" type="text" autocomplete="shipping postal-code" data-friendly-error="Shipping postal code is required" />
	<div class="invalid-feedback">
		Zip / Postal Code is required
	</div>
	<div class="invalid-feedback feedback-icon">
		<i class="fa fa-times"></i>
	</div>
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
</div>

<div class="form-group position-relative">
	<label for="shipping-phone">Phone #</label>
	<input class="form-control" name="shipping-phone" type="text" autocomplete="shipping tel" />
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
</div>

	<div class="custom-control custom-checkbox" <%= hide_non_registered %> id="shipping-save-wrapper">
			<input type="checkbox" class="custom-control-input" name="shipping-save" id="shipping-save">
			<label class="custom-control-label" for="shipping-save">Save this shipping address to my account</label>
	</div>
<input name="shipping-status" id="shipping-status" type="hidden" value="">
</div><!-- end shipping address form -->
</div><!-- end hide content for adding on products -->
</div><!-- end shipping card body-->
</div><!-- end shipping main card -->
</section><!-- end shipping address section -->


<% if var_grandtotal + session("temp_shipping") = 0 then
	hide_payment_section = "style=""display:none"""
end if 

if request.querystring("type") <> "paypal" and request.querystring("type") <> "afterpay" then ' Only display billing if checking out with credit card
%>
<section class="billing-information" id="billing-information" <%= hide_payment_section %>> 
<div class="card mb-5">
<div class="card-header"><h5>Payment method</h5></div>
<div class="card-body">
		<button class="btn btn-sm btn-info add-new-billing-button" style="display:none" type="button">Add a new credit card</button>
		<button class="btn btn-sm btn-outline-danger mb-4" style="display:none" type="button" id="cancel-billing-add">Cancel Add / Update</button>
		
	<div style="display:none" id="cim_billing_addresses">
	<!--#include virtual="/checkout/cim_getbilling_addresses.asp"-->
</div>



<div class="billing-address-form AddressesForm">
<div class="billing-input-fields">

<div class="custom-control custom-checkbox" id="shipping-same-billing-wrapper" <%= hide_section_addons %>>
	<input type="checkbox" class="custom-control-input" name="shipping-same-billing" id="shipping-same-billing">
	<label class="custom-control-label" for="shipping-same-billing">Billing address is the same as my shipping address</label>
</div>

<div class="mb-3" id="credit_card_inputs">
		<div class="text-secondary mt-1" style="font-size:2em">
				<i class="fa fa-cc-visa"></i>
				<i class="fa fa-cc-mastercard"></i>
				<i class="fa fa-cc-amex"></i>
				<i class="fa fa-cc-discover"></i>
			</div>

<div class="form-group position-relative">
	<label for="cardNumber">Card number <span class="text-danger">*</span></label>
	<input class="form-control" required type="tel" name="card_number" id="cardNumber" placeholder="&#8226;&#8226;&#8226;&#8226;  &#8226;&#8226;&#8226;&#8226;  &#8226;&#8226;&#8226;&#8226;  &#8226;&#8226;&#8226;&#8226;" data-validation="length alphanumeric"  data-validation-length="12-19" data-validation-allowing=" " autocomplete="cc-number" data-friendly-error="Credit card # is required" />
	<div class="invalid-feedback">
			Valid credit card # is required
		</div>
		<div class="valid-feedback feedback-icon">
			<i class="fa fa-check"></i>
		</div>
		<div class="invalid-feedback feedback-icon">
			<i class="fa fa-times"></i>
		</div>
</div>
		
<div class="form-group position-relative">
	<label for="creditCardMonth">Exp month <span class="text-danger">*</span></label>
	<select class="form-control" required name="billing-month" id="creditCardMonth" name="billing-month"  autocomplete="cc-exp-month" data-friendly-error="Expiration month is required" >
			<option value="">Select month</option>
			<option value="01">01 - January</option>
			<option value="02">02 - February</option>
			<option value="03">03 - March</option>
			<option value="04">04 - April</option>
			<option value="05">05 - May</option>
			<option value="06">06 - June</option>
			<option value="07">07 - July</option>
			<option value="08">08 - August</option>
			<option value="09">09 - September</option>
			<option value="10">10 - October</option>
			<option value="11">11 - November</option>
			<option value="12">12 - December</option>
		</select>
		<div class="invalid-feedback">
				Month is required
			</div>
			<div class="valid-feedback feedback-icon">
				<i class="fa fa-check"></i>
			</div>
			<div class="invalid-feedback feedback-icon">
				<i class="fa fa-times"></i>
			</div>
	</div>
		
<div class="form-group position-relative">
	<label for="creditCardYear">Exp year <span class="text-danger">*</span></label>
	<select class="form-control" required name="billing-year" id="creditCardYear" autocomplete="cc-exp-year" data-friendly-error="Expiration year is required">
			<option value="">Select year</option>
		<% for i = 0 to 10 %>
			<option value="<%= year(now) + i %>"><%= year(now) + i %></option>
		<% next %>
	</select>
	<div class="invalid-feedback">
			Year is required
		</div>
		<div class="valid-feedback feedback-icon">
			<i class="fa fa-check"></i>
		</div>
		<div class="invalid-feedback feedback-icon">
			<i class="fa fa-times"></i>
		</div>
</div>

<div class="form-group position-relative">
	<label for="cvv2">Security code (3-4 digit number on back of card) <span class="text-danger">*</span></label>
	<input class="form-control" type="tel" name="cvv2" id="security-code">
	<div class="invalid-feedback">
			Security code is required
		</div>
		<div class="valid-feedback feedback-icon">
			<i class="fa fa-check"></i>
		</div>
		<div class="invalid-feedback feedback-icon">
			<i class="fa fa-times"></i>
		</div>
</div>
</div><!-- credit card input fields -->
<div class="form-group position-relative">
	<label for="billing-first" class="control-label">First name <span class="text-danger">*</span></label>
	<input class="form-control" required name="billing-first" id="billing-first" type="text" autocomplete="billing given-name" data-friendly-error="Billing address: First name is required" />
	<div class="invalid-feedback">
		First name is required
	</div>
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
	<div class="invalid-feedback feedback-icon">
		<i class="fa fa-times"></i>
	</div>
</div>

<div class="form-group position-relative">
	<label for="billing-last" class="control-label">Last name <span class="text-danger">*</span></label>
	<input class="form-control" required name="billing-last" id="billing-last" type="text" autocomplete="billing family-name" data-friendly-error="Billing address: Last name is required" />
	<div class="invalid-feedback">
		Last name is required
	</div>
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
	<div class="invalid-feedback feedback-icon">
		<i class="fa fa-times"></i>
	</div>
</div>

<div class="form-group position-relative">
	<label for="billing-address">Address (Line 1)<span class="text-danger">*</span></label>
	<input class="form-control" required name="billing-address" id="billing-address" type="text" autocomplete="billing address-line1" data-friendly-error="Billing address is required" />
	<div class="invalid-feedback">
		Address is required
	</div>
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
	<div class="invalid-feedback feedback-icon">
		<i class="fa fa-times"></i>
	</div>
</div>

<div class="form-group position-relative">
	<label for="billing-address2">Apt #, Dorm, Suite&nbsp;&nbsp;</label>
	<input class="form-control" name="billing-address2" id="billing-address2" type="text" autocomplete="billing address-line2" />
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
</div>

<div class="form-group position-relative">
	<label for="billing-city">City <span class="text-danger">*</span></label>
	<input class="form-control" required name="billing-city" id="billing-city" type="text" autocomplete="billing address-level2" data-friendly-error="Billing city is required" />
	<div class="invalid-feedback">
		City is required
	</div>
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
	<div class="invalid-feedback feedback-icon">
		<i class="fa fa-times"></i>
	</div>
</div>

<div class="form-group position-relative">                     
	<label for="billing-country">Country <span class="text-danger">*</span></label>
	<select class="form-control" required name="billing-country"  id="billing-country" autocomplete="billing country" data-friendly-error="Billing country is required" >
	<option value="USA" selected>USA</option>
			<% 
	While NOT rsGetCountrySelect.EOF 
	%>
			<option value="<%=(rsGetCountrySelect.Fields.Item("Country").Value)%>"><%=(rsGetCountrySelect.Fields.Item("Country").Value)%></option>
			<% 
	rsGetCountrySelect.MoveNext()
	Wend
	rsGetCountrySelect.Requery


	%>
		</select>
		<div class="invalid-feedback">
			Country is required
		</div>
		<div class="valid-feedback feedback-icon">
			<i class="fa fa-check"></i>
		</div>
		<div class="invalid-feedback feedback-icon">
			<i class="fa fa-times"></i>
		</div>
</div>

<div class="form-group position-relative billing-state <%= hide_state %>">
	<label for="billing-state">State (USA)<span class="text-danger">*</span></label>
	<select class="form-control" name="billing-state" id="billing-state" autocomplete="billing address-level1" data-friendly-error="Billing state is required">
	<!--#include file="includes/inc_states_select.asp"-->
		</select>
		<div class="invalid-feedback">
			State is required
		</div>
		<div class="valid-feedback feedback-icon">
			<i class="fa fa-check"></i>
		</div>
		<div class="invalid-feedback feedback-icon">
			<i class="fa fa-times"></i>
		</div>
</div>

<div class="form-group position-relative billing-province-canada <%= hide_canada %>">
	<label for="billing-province-canada">Province <span class="text-danger">*</span></label>
	<select class="form-control" name="billing-province-canada" id="billing-province-canada" autocomplete="billing address-level2" data-friendly-error="Billing province is required">
	<!--#include file="includes/inc_province_canada_select.asp"-->
		</select>
		<div class="invalid-feedback">
			Province is required
		</div>
		<div class="valid-feedback feedback-icon">
			<i class="fa fa-check"></i>
		</div>
		<div class="invalid-feedback feedback-icon">
			<i class="fa fa-times"></i>
		</div>
</div>

<div class="form-group position-relative billing-province <%= hide_province %>">
	<label for="billing-province">Province / State</label>
	<input class="form-control" name="billing-province" id="billing-province" type="text" autocomplete="billing address-level2" data-friendly-error="Billing province/state is required" />
	<div class="invalid-feedback">
		Province / State is required
	</div>
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
	<div class="invalid-feedback feedback-icon">
		<i class="fa fa-times"></i>
	</div>
</div>

<div class="form-group position-relative">
	<label for="billing-zip"><span class="hide_usa_zip <%= hide_state %>">Zip code</span><span class="hide_inter_zip <%= hide_inter_zip %>">Postal code</span> <span class="text-danger">*</span></label>
	<input class="form-control" name="billing-zip" id="billing-zip" type="text" autocomplete="billing postal-code" data-friendly-error="Billing postal code is required" />
	<div class="invalid-feedback">
		Zip / Postal Code is required
	</div>
	<div class="invalid-feedback feedback-icon">
		<i class="fa fa-times"></i>
	</div>
	<div class="valid-feedback feedback-icon">
		<i class="fa fa-check"></i>
	</div>
</div>


<div class="custom-control custom-checkbox" id="card-save-wrapper" <%= hide_non_registered %>>
	<input type="checkbox" class="custom-control-input" name="card-save" id="card-save">
	<label class="custom-control-label" for="card-save">Save this credit card to my account</label>
</div>

<button class="btn btn-sm btn-info" id="btn-save-credit-card" data-id="" style="display:none" type="button">Save my new credit card # <i class="fa fa-spinner fa-spin fa-lg ml-3" style="display:none" id="spinner-update-billing"></i></button>
<div class="alert alert-danger mt-2" id="msg-update-billing" style="display:none"></div>

</div>
<input name="billing-status" id="billing-status" type="hidden" value="">

<div <%= hide_registered %> <%= hide_section_addons %>><!-- hide different payment methods if registered -->
	<div class="billing-input-fields">
		<h4>OR</h4>
	</div>
	<div class="custom-control custom-checkbox">
		<input type="checkbox" class="custom-control-input" name="cash" id="cash" value="on">
		<label class="custom-control-label" for="cash">Money order or cash</label>
	</div>
</div><!-- hide different payment methods if registered -->
</div><!-- end billing address form -->
</div><!-- end billing card body -->
</div><!-- end billing main card -->
</section><!-- end billing address section -->
<% end if ' if checkout type is not paypal %>
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->

<section class="shipping-options" <%= hide_section_addons %>> 
	<div class="card mb-5">
<div class="card-header"><h5>Shipping options</h5></div>
<div class="card-body">
<%
if preorder_shipping_notice = "yes" then
%>
<div class="alert alert-warning font-weight-bold">
		Your order contains CUSTOM ORDERED items.
<br/>		Your ENTIRE ORDER will be held until the custom piece arrives to ship to you.
</div>
<% 
else
%>
<div class="alert alert-info font-weight-bold">
<%
			' AMANDA ADDED - special code to give more info on when package will ship out
			If WeekDayName(WeekDay(date())) = "Saturday" OR WeekDayName(WeekDay(date())) = "Sunday" then
			Response.Write "Your order will SHIP OUT MONDAY"
			end if
			
			If Time() > "08:00:00 AM" AND WeekDayName(WeekDay(date())) <> "Saturday" AND WeekDayName(WeekDay(date())) <> "Sunday" AND WeekDayName(WeekDay(date())) <> "Friday" then
			Response.Write "Your order will SHIP OUT TOMORROW"
			end if
			
			If Time() > "08:00:00 AM" AND WeekDayName(WeekDay(date())) = "Friday" then
			Response.Write "Your order will SHIP OUT MONDAY"
			end if
			
			If Time() < "08:00:00 AM" AND WeekDayName(WeekDay(date())) <> "Saturday" AND WeekDayName(WeekDay(date())) <> "Sunday" then
			Response.Write "Your order will SHIP OUT TODAY"
			end if
%>
</div>
<%		
end if ' if a custom order is found in the order

if request.form("shipping-country") = "" and session("shipping-display") = "" then 
	var_hide = "style=""display:none"""
end if

'response.write "Subtotal after discounts: " & var_subtotal_after_discounts
%>
<div class="customs-notice" style="display:none">
	<div class="alert alert-info">
	For shipping outside of the US, customs rules and regulations may apply and can cause delays in delivery that are outside the control of the USA Postal Service. 
	</div>
</div>
<div class="alert alert-primary p-1" style="display:none" id="shipping-loading"><i class="fa fa-spinner fa-2x fa-spin"></i> Loading shipping options...</div>
<div class="shipping-section" <%= var_hide %>>
<div class="btn-group btn-group-toggle flex-wrap w-100" data-toggle="buttons" id="ajax-shipping-options">
</div><!-- button group -->	
</div><!-- no display if scripts disabled -->
</div><!-- end shipping options card body -->
</div><!-- end shipping options main card -->
</section><!-- end shipping options section -->

<% if CustID_Cookie = 0 then ' show section if not registered  %>
<section class="create-password AddressesForm" <%= hide_section_addons %>> 
	<div class="card mb-5">
		<div class="card-header">
			<h5>Create password & account (optional)</h5>
		</div>
		<div class="card-body">
				<div id="duplicate_account" class="alert alert-danger" style="display:none"></div>
				<div class="form-group">
					<label for="pass_confirmation_checkout">Password</label>
					<input class="form-control" type="password" name="password_confirmation" id="pass_confirmation_checkout" data-validation-error-msg="Password is required" class="validate-ignore" autocomplete="off">
				</div>
				<div class="form-group">
					<label for="password_checkout">Re-type password</label>
					<input class="form-control" type="password" name="password" id="password_checkout" data-validation="confirmation" data-validation-error-msg="Passwords don't match" class="validate-ignore" autocomplete="off">
				</div>
				<div class="custom-control custom-checkbox mt-2">
					<input type="checkbox" class="custom-control-input" name="save-all" id="save-all" value="on">
					<label class="custom-control-label" for="save-all">Save all my shipping & billing information to my new account</label>
				</div>
		</div><!-- password card body -->
	</div><!-- password card -->
</section><!-- end create password section -->
<% end if ' show section if not registered  %>
<section>
	<div class="card  mb-5 mb-lg-0">
		<div class="card-header"><h5>Review Your Cart</h5></div>
		<div class="card-body">
			
<div class="container-fluid">
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
	<div class="row">
                 <div class="col-auto col-xl-auto">
				  <% If Instr(rs_getCart.Fields.Item("title").Value, "Digital gift certificate") > 0 Then
					product_link = "gift-certificate.asp"
				  else
					product_link = "productdetails.asp?ProductID=" & rs_getCart.Fields.Item("ProductID").Value
				  end if
				  %>
				  <a href="<%= product_link %>"><div class="position-relative"><img  src="https://s3.amazonaws.com/bodyartforms-products/<%=(rs_getCart.Fields.Item("picture").Value)%>" alt="Product photo">

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
				 <div class="col col-lg-9 col-xl-5 small pl-0">	
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
					<strong>Your specs:</strong> <span class="spectext<%= rs_getCart.Fields.Item("cart_id").Value %>"><%= rs_getCart.Fields.Item("cart_preorderNotes").Value %></span>
				
				<% else ' show gift certificate information 
					certificate_array =split(rs_getCart.Fields.Item("cart_preorderNotes").Value,"{}")				
				%>
				<span class="font-weight-bold">Recipient's name:</span> <%= certificate_array(3) %>
				<span class="font-weight-bold">Recipient's e-mail:</span> <%= certificate_array(0) %>
				<span class="font-weight-bold">Your name:</span> <%= certificate_array(1) %>
				<span class="font-weight-bold">Your message:</span> <%= certificate_array(2) %>
				<%	end if ' detect whether preorder or gift cert %>
			<% end if %>
		
			<% if rs_getCart.Fields.Item("cart_qty").Value <= rs_getCart.Fields.Item("qty").Value then %>
			<% if rs_getCart.Fields.Item("customorder").Value = "yes" then 
			preorder_in_order = "yes"
			%>
					<span class="d-inline-block my-1 bg-info text-white p-2">
						<%= rs_getCart.Fields.Item("preorder_timeframes").Value %> to receive
					</span>	
			<% else %>
			
			<% end if %>
			<% end if %>
		
      </div><!-- end col / item information -->
			<div class="col-12 col-lg-12 col-xl pt-2 py-xl-0">

	<div class="d-inline d-xl-block">	
	Qty: <%= rs_getCart.Fields.Item("cart_qty").Value %>

		@ 	  
					<span class="mr-1" data-price="<%= FormatNumber(var_itemPrice, -1, -2, -2, -2) %>"><%= exchange_symbol %><%= FormatNumber(var_itemPrice, -1, -2, -2, -2) %></span><span  class="mr-3"><%= qty_pair_text %></span>
					<%
					if FormatNumber(var_itemPrice, -1, -2, -2, -2) < FormatNumber(rs_getCart.Fields.Item("price").Value * exchange_rate, -1, -2, -2, -2) then
					%>
					<strike class="mr-1" data-price="<%= FormatNumber(rs_getCart.Fields.Item("price").Value * exchange_rate, -1, -2, -2, -2) %>"><%= exchange_symbol %><%= FormatNumber(rs_getCart.Fields.Item("price").Value * exchange_rate, -1, -2, -2, -2) %></strike>
					<% end if %>					                


</div><!-- end qty display -->
<div class="d-inline d-xl-block">
			<span class="font-weight-bold line_item_total_<%= rs_getCart.Fields.Item("ProductDetailID").Value %>" data-price="<%= FormatNumber(var_lineTotal, -1, -2, -2, -2) %>"><%= exchange_symbol %><%= FormatNumber(var_lineTotal, -1, -2, -2, -2) %></span>
			<span class="font-weight-bold ml-1">total</span>
		<% if (rs_getCart.Fields.Item("SaleDiscount").Value <> 0 or Session("CouponPercentage") <> "" OR Session("Preferred") = "yes") AND var_giftcert <> "yes" then  ' only display if the item is cheaper than retail 
%>

		<% if rs_getCart.Fields.Item("SaleExempt").Value = 1 AND (Session("Preferred") = "yes" and Session("CouponPercentage") <> "")then %>

						<span class="d-inline-block badge badge-warning p-1 rounded-0">Coupon exempt</span>
						<% 	end if
		 end if%>
			</div><!-- end line total block -->
	</div><!-- end col /  totals and qty box -->
</div><!-- end row -->
<hr>
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
</div><!-- entire cart container -->
<div class="small" <%= hide_section_addons %>>
	<%= var_free_items %>
</div>
		</div><!-- review cart BODY-->
	</div><!-- review cart MAIN CARD -->
</section><!-- review cart section --> 

<input type="hidden" name="weight" id="weight" value="<%= session("weight") %>">

</div><!-- end cart items container -->
</div><!-- end items column-->
	<div class="col-12 col-lg-4 col-break1600 col-break1900 m-0 p-0">
		<div class="sticky-top" style="z-index:100">
			<div class="card bg-light mb-2">
				<div class="card-body text-left py-2">
								<div class="row">	
									<div class="col-7">Subtotal</div><div class="col-5">$<span class="cart_subtotal"><%= FormatNumber((var_subtotal), -1, -2, -2, -2) %></span></div>
								</div>		
								<% if Session("CouponCode") <> "" then %>
								<div class="row">
									<div class="col-7">Coupon</div><div class="col-5">- $<span class="cart_coupon-amt"><%= FormatNumber(var_couponTotal, -1, -2, -2, -2) %></span></div>
								</div>
								<% 
								end if 
								%>
								<% if Request.Cookies("ID") <> "" then 
								%>
								 <% if TotalSpent > 275 AND Session("CouponCode") = "" then %>
									<div class="row">
									<div class="col-7">Your 10% discount</div><div class="col-5">- <span class="cart_prefferred_discount"><%= FormatCurrency(total_preferred_discount, -1, -2, -2, -2) %></span>
									</div></div>
								<% 
								end if ' if preferred customer 
								%>
								<%
								 end if ' if customer is logged in
								
								%>
								<% 
								if Session("GiftCertAmount") <> 0 then 
								%>
									<div id="row_gift_cert">
										<div class="row">
									<div class="col-7">Gift certificate</div><div class="col-5">- <span id="cart_gift-cert"><%= FormatCurrency(Session("GiftCertAmount"), -1, -2, -2, -2) %></span></div>
								</div>
									</div>
								<% ' if there is a gift certificate found
								end if 
								%>
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
										<div class="col-7"><span class="cart_sales-tax-state"></span>Tax</div><div class="col-5 cart_sales-tax"><%= var_salesTax %></div>
									</div>
								</div><!-- end card body -->
						<div class="card-footer">
								
									<h4>TOTAL <% if strcountryName <> "US" then %> (USD)<% end if %>$<span class="cart_grand-total"><%= FormatNumber(var_grandtotal, -1, -2, -2, -2) %></span></h4>
							<div class="row_convert_total" style="display:none">
								<div class="alert alert-success p-2">
									<div><h6><img class="mr-2" style="width:20px;height:20px" id="currency-icon" src="/images/icons/<%= currency_img %>">ESTIMATE <span class="exchange-price"><span class="currency-type"></span> <span class="convert-total convert-price" data-price=""></span></span></h6></div>
										<span class="exchange-price"><span class="currency-type bold"></span> <span class="convert-total convert-price bold" data-price=""></span> is a close estimate</span>. The total billed will be for <span class="bold">$<span class="cart_grand-total"><%= FormatNumber(var_grandtotal, -1, -2, -2, -2) %></span> in US Dollars</span> and your bank will convert to the most current exchange rate.
								</div>
						</div>
						<div class="alert alert-danger stock-error" style="display:none"></div>				
							<% if var_only_gift_cert <> "yes" then %>
							<div class="alert alert-danger submit_disabled" style="display:none"></div><!-- alert for that displays if a shipping type has not been selected -->
							<% end if %>
					<% if request.querystring("type") = "card" then %>
						<% If toggle_checkout_cards = true Then %>
							<button class="btn btn-lg btn-primary btn-block checkout_button place_order" style="display: none" type="submit" form="checkout_form" name="place_order">PLACE ORDER</button>
						<% else %>
							<div class="alert alert-danger">We're sorry, but our <b>credit card</b> checkout is temporarily unavailable. As soon as our payment processor comes back online, we will accept orders again. Please check back later.</div>
						<% end if %>
					<% end if %>
					<% if request.querystring("type") = "paypal" then %>
						<% If toggle_checkout_paypal = true Then %>
							<button class="btn btn-lg btn-warning btn-block checkout_button place_order checkout_paypal" style="display:none" type="submit" form="checkout_form" name="place_order">CONTINUE TO <img class="ml-1" style="height:25px" src="/images/paypal.png" /></button>
							<input type="hidden" name="paypal" value="on">
						<% else %>
							<div class="alert alert-danger">We're sorry, but our <b>PayPal</b> checkout is temporarily unavailable. As soon as PayPal comes back online, we will accept orders again. Please check back later.</div>
						<% end if %>						
					<% end if %>
					<%
					' === only show afterpay option to USA customers
					if (request.cookies("currency") = "" OR request.cookies("currency") = "USD") AND request.querystring("type") = "afterpay" then
						afterpay_display = ""
					else
						afterpay_display = "display:none"
					end if
					%>
					<!--
					<div id="REMOVE-GO-LIVE" style="display:none">
					<div class="afterpay_option" style="<%= afterpay_display %>">
						<button class="btn btn-lg btn-primary btn-block checkout_button place_order checkout_afterpay" style="display:none" type="submit" form="checkout_form" name="place_order">PAY NOW WITH <img class="img-fluid d-inline w-50" src="/images/afterpay-white-logo.png"/></button>
						<span class="d-none"><span class="afterpay-widget-nonactive afterpay-widget"></span></span>
						<% if request.querystring("type") = "afterpay" then %> 
						<input type="hidden" name="afterpay" value="on">
						<% end if %>
					</div>
					</div>
					-->
					<div class="processing-message" style="display:none"></div>			
						
						
					<% if preorder_in_order = "yes" then %>
					<div class="alert alert-warning p-1 my-2 small font-weight-bold">
						Custom order cancellations will have a 15% restocking fee
					</div>
					<% end if %>	
				</div><!-- end card footer for totals -->
				</div><!-- end card for totals -->
				<% if request.cookies("OrderAddonsActive") = "" then %>
				<div class="custom-control custom-checkbox mt-3">
					<input type="checkbox" class="custom-control-input  event-newsletter" name="newsletter-signup" id="checkout-newsletter-signup" value="">
					<label class="custom-control-label font-weight-bold" for="checkout-newsletter-signup">STAY CONNECTED <i class="fa fa-paper-plane ml-2"></i></label>
					<div class="small">
						Sign up for our newsletter and get notified anytime we run sales or special events
				</div>
				</div>
				<div class="custom-control custom-checkbox mt-2">
					<input type="checkbox" class="custom-control-input" name="conserve-plastic" id="conserve" value="CONSERVE PLASTIC BAGS<br>">
					<label class="custom-control-label font-weight-bold" for="conserve">CONSERVE PLASTIC <i class="fa fa-recycle ml-2"></i></label>
					<div class="small">
							Check here to have us  put as many items as safely possible into one baggie and seal it all together. The only drawback to this is if you want to return one item out of the bag. You will have to return the entire  sealed bag, because breaking the seal voids a return.
					</div>
				</div>
				<div class="custom-control custom-checkbox mt-2">
					<input type="checkbox" class="custom-control-input" name="gift" id="gift" value="GIFT ORDER<br>">
					<label class="custom-control-label small" for="gift">Gift order? Check here to not print the prices on the invoice </label>
				</div>
				<div class="form-group mt-3">
					<label for="comments">Order comments:</label>
					<textarea class="form-control" name="comments" id="comments" rows="4"></textarea>
				</div>
			<!-- display bottom section for mobile only -->
			
			<div class="d-lg-none">
				<div class="alert alert-danger stock-error" style="display:none"></div>
					<% if request.querystring("type") = "card" then %>
						<% If toggle_checkout_cards = true Then %>
							<button class="btn btn-lg btn-primary btn-block checkout_button place_order" style="display: none" type="submit" form="checkout_form" name="place_order">PLACE ORDER</button>
						<% else %>
							<div class="alert alert-danger">We're sorry, but our <b>credit card</b> checkout is temporarily unavailable. As soon as our payment processor comes back online, we will accept orders again. Please check back later.</div>
						<% end if %>
					<% end if %>
					<% if request.querystring("type") = "paypal" then %>
						<% If toggle_checkout_paypal = true Then %>
							<button type="submit" form="checkout_form" name="place_order" class="btn btn-lg btn-warning btn-block checkout_button place_order checkout_paypal mt-4" style="display: none">CONTINUE TO <img class="ml-1" style="height:25px" src="/images/paypal.png" /></button>
						<% else %>
							<div class="alert alert-danger">We're sorry, but our <b>PayPal</b> checkout is temporarily unavailable. As soon as PayPal comes back online, we will accept orders again. Please check back later.</div>
						<% end if %>		
					<% end if %>
					<!--
					<div id="REMOVE-GO-LIVE" style="display:none">
						<div class="afterpay_option" style="<%= afterpay_display %>">
							<button class="btn btn-lg btn-primary btn-block checkout_button place_order checkout_afterpay" style="display:none" type="submit" form="checkout_form" name="place_order">PAY NOW WITH <img class="img-fluid d-inline w-50" src="/images/afterpay-white-logo.png"/></button>
						</div>
						</div>
					<-->
					<div class="processing-message" style="display:none"></div>
				</div><!-- display bottom section for mobile only -->
				<% end if ' OrderAddonsActive is null %>
				<% end if ' if order or account is not flagged %>

</div><!-- sticky top -->
</div><!-- totals column -->
</div><!-- entire cart row -->
</div><!-- entire cart container -->

<% if var_only_gift_cert = "yes" then %>
	<input type="hidden" id="gift_cert_only" value="yes">
<% end if %>
<%
 end if ' Show if cart is NOT empty
%>
</form>
<div id="load_temps"></div>
<input type="hidden" id="timestamp" value="<%= now() %>">

<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript">
	$('#cim_shipping_addresses, #cim_billing_addresses').show();
</script>
<% if CustID_Cookie <> 0 then %>
<script type="text/javascript">
	// Disable all form fields by default
	$('.shipping-address-form input, .shipping-address-form select, .billing-address-form input, .billing-address-form select').attr('disabled', true);

</script>
<% end if %>
<% if CustID_Cookie <> 0 and var_no_ship_addresses = "true" then 
%>
<script type="text/javascript">
	$('.shipping-address-form input, .shipping-address-form select').attr('disabled', false);
	$('.billing-address-form input, .billing-address-form select').attr('disabled', false);
		$('#billing-status').val("add");
		$('#shipping-status').val("add");
		$('#cc_logos').hide();
</script>
<% end if %>

<% if CustID_Cookie <> 0 and var_no_ship_addresses = "false" then 
%>
	<script type="text/javascript">
	$('.shipping-address-form, .billing-address-form').hide();
	$('.add-new-shipping-button, .add-new-billing-button').show();
	</script>
<% end if %>
<%	if request.cookies("OrderAddonsActive") <> "" then	%>
<script type="text/javascript">
	$('#e-mail, #shipping-first-checkout, #shipping-last-checkout, #shipping-address, #shipping-city, #shipping-state, #shipping-zip').prop('required', false);
</script>
<%	end if 	%>
<script type="text/javascript" src="/js-pages/toggle_required_billing.js"></script>
<script type="text/javascript" src="/js-pages/currency-exchange.min.js?v=050619"></script>
<% if (session("exchange-rate") = "" OR session("exchange-currency") <> request.cookies("currency")) AND request.cookies("currency") <> "" AND request.cookies("currency") <> "USD" then %>
<script>
		// Get currency conversions on page load
		updateCurrency();
</script>
<% end if %>
<script type="text/javascript" src="/js-pages/cart_update_totals.min.js?v=102721"></script>
<script type="text/javascript" src="/js-pages/cart.min.js?v=050319" async></script>
<script type="text/javascript" src="/js-pages/checkout.min.js?v=111121"></script>
<!-- Start Afterpay Javascript -->
<!--
<script src="https://portal.sandbox.afterpay.com/afterpay.js" async></script>-->
<!--
<script type = "text/javascript" src="https://static-us.afterpay.com/javascript/present-afterpay.js"></script>-->
<!--
<script type="text/javascript" src="/js-pages/afterpay-widget.js?v=020420" ></script>-->

<%
Set rsToggles = Nothing
%>