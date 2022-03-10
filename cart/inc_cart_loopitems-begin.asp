<% 
' Cart repeat and display items
var_subtotal = 0
var_couponTotal = 0
var_couponSubtotal = 0
var_anodizationSubTotal = 0
var_totalweight = 0
' added 6/27/18 to fix a 500 error /cart/inc_display_estimated_shipping.asp Line 59 Incorrect syntax near '>'.. .
session("weight")  = 0
var_item_count = 0
var_total_discounts = 0
var_discount_preferred = 0
total_preferred_discount = 0
shipping_cost = 0
var_other_items = 0
var_totalvalue_certs_incart = 0
exempt_item_in_cart = ""
var_autoclavable = 0
var_sterilization_added = 0
var_anodization_added = 0
credit_now = 0 ' added May 2016 to fix PayPal issue of doubling up on credits in email confirmation
str_autoclave_items = ""


if session("storeCredit_used") = "" then
	session("storeCredit_used") = 0
end if

if session("credit_now") = "" then
	session("credit_now") = 0
end if

rs_getCart.ReQuery()
While Not rs_getCart.EOF

var_item_count = var_item_count + 1

	var_itemPrice = 0
	var_itemPrice_USdollars = 0

	var_itemPrice =  FormatNumber(rs_getCart.Fields.Item("price").Value * exchange_rate, -1, -2, -2, -2)
	var_itemPrice_USdollars =  FormatNumber(rs_getCart.Fields.Item("price").Value, -1, -2, -2, -2)

	if (rs_getCart.Fields.Item("SaleDiscount").Value > 0 AND rs_getCart.Fields.Item("secret_sale").Value = 0) OR (rs_getCart.Fields.Item("secret_sale").Value = 1 AND session("secret_sale") = "yes") then
		var_itemPrice = FormatNumber(var_itemPrice - (rs_getCart.Fields.Item("price").Value * exchange_rate * rs_getCart.Fields.Item("SaleDiscount").Value/100), -1, -2, -2, -2)
		var_itemPrice_USdollars = FormatNumber(var_itemPrice_USdollars - (rs_getCart.Fields.Item("price").Value * rs_getCart.Fields.Item("SaleDiscount").Value/100), -1, -2, -2, -2)
	end if
	
	'Weight in DB is a not null field that will be 0 or a value
	var_totalweight = var_totalweight + (rs_getCart.Fields.Item("weight").Value * rs_getCart.Fields.Item("cart_qty").Value)
	session("weight") = var_totalweight
	
if rs_getCart.Fields.Item("SaleExempt").Value = 1 then 
		var_saleExempt = "yes"
else
		var_saleExempt = "no"
end if


' flag which items are actually eligible for a coupon discount, SaleExempt, gift certs, and brands
if var_saleExempt = "yes" then

	exempt_item_in_cart = "yes"

	if TotalSpent > 275 and CustID_Cookie <> "" then
		if Session("CouponCode") = "" then
			var_discount_preferred = var_itemPrice_USdollars * .10 * rs_getCart.Fields.Item("cart_qty").Value
		else
			var_discount_preferred = 0
			var_discount_coupon = var_itemPrice_USdollars * .10
			var_line_coupon = var_itemPrice_USdollars * .10 * rs_getCart.Fields.Item("cart_qty").Value
		end if
		var_creditType = "Discount"
	else
		var_discount_coupon = 0
		var_line_coupon = 0
		var_creditType = "Coupon"
	end if
else ' if not sales exempt
	if Session("CouponPercentage") <> "" then ' if a coupon is active
		var_discount_coupon = (var_itemPrice_USdollars * (Session("CouponPercentage") / 100))
		var_line_coupon = (var_itemPrice_USdollars * (Session("CouponPercentage") / 100)) * rs_getCart.Fields.Item("cart_qty").Value
		var_creditType = "Coupon"
	
		'Set discount to $0 if it's a brand coupon and doesn't match
		if session("brand_coupon") <> "" OR session("brand_coupon") <> "None" then
			if Instr(1, rs_getCart.Fields.Item("brandname").Value, session("brand_coupon")) = 0 Then
				var_discount_coupon = 0
				var_line_coupon = 0
				var_creditType = "Coupon"
		
				if TotalSpent > 275 and CustID_Cookie <> "" then
					var_discount_preferred = (var_itemPrice_USdollars * .10) * rs_getCart.Fields.Item("cart_qty").Value
					var_creditType = "Discount"
				else
					var_discount_coupon = var_itemPrice_USdollars
					var_line_coupon = 0
					var_creditType = "Coupon"
				end if	
			end if
		end if
	else
		if TotalSpent > 275 and CustID_Cookie <> "" then
			var_discount_preferred = (var_itemPrice_USdollars * .10) * rs_getCart.Fields.Item("cart_qty").Value
			var_creditType = "Discount"
		else
			var_discount_coupon = 0
			var_line_coupon = 0
			var_creditType = "Coupon"
		end if
	end if
end if

'Set gift certificate purchases to no discount
var_giftcert = "no" ' default state
'If Instr(lcase(rs_getCart.Fields.Item("title").Value), lcase("Digital gift certificate")) > 0 Then
if rs_getCart("ProductID") = 2424 then
'		response.write "gift cert found"
		var_discount_preferred = 0
		var_discount_coupon = 0
		var_line_coupon = 0
		var_giftcert = "yes"
		session("var_giftcert") = "yes"
		var_totalvalue_certs_incart = var_totalvalue_certs_incart + (rs_getCart.Fields.Item("price").Value * rs_getCart.Fields.Item("cart_qty").Value)
else ' detect items in cart OTHER than gift certs (this is for showing free items on the cart page)
	var_other_items = 1
	session("var_other_items") = 1
'	response.write "gift cert NOT found"
end if


if var_other_items <> 1 then ' if only gift certs are found erase all free gift cookies
	response.cookies("oringsid") = ""
	response.cookies("gaugecard") = "no"
	response.cookies("stickerid") = ""
	response.cookies("freegift1id") = ""
	response.cookies("freegift2id") = ""
	response.cookies("freegift3id") = ""
	response.cookies("freegift4id") = ""
	response.cookies("freegift5id") = ""
end if

if rs_getCart.Fields.Item("customorder").Value = "yes" then 
	preorder_shipping_notice = "yes"
end if

'Detect if there are ANY autoclavable items in the order. default "0" is set above, 1 will act as a "yes"
if rs_getCart.Fields.Item("autoclavable").Value = "1" then
	var_autoclavable = 1
end if

' Detect if sterilization service has already been added to order tracking the Product DETAIL ID to see if it's added since this is unlikely to change
if rs_getCart.Fields.Item("cart_detailID").Value = 34356 then
	var_sterilization_added = 1
	var_cart_id_autoclave = rs_getCart.Fields.Item("cart_id").Value
end if

if CLng(rs_getCart.Fields.Item("anodID").Value) > 0 then
	var_anodization_added = 1
end if
'======= BUILD AUTOCLAVE STRING TO BE DISPLAYED ON CART.ASP PAGE =========================
if rs_getCart("autoclavable") = 1 then
	str_autoclave_items = "<li>" & rs_getCart("title") & "</li>" & str_autoclave_items
end if

var_anodizationPrice = 0
If var_anodization_added = 1 Then
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM TBL_Anodization_Colors_Pricing WHERE anodID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",200,1,70,rs_getCart.Fields.Item("anodID").Value))
	Set rsAnodization = objCmd.Execute()
	If Not rsAnodization.EOF Then 
		var_anodizationBasePrice = rsAnodization("base_price")
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT SUM(cart_qty) as count FROM tbl_carts WHERE anodID > 0 AND (" & var_db_field & " = ?) AND (cart_save_for_later = 0 OR cart_save_for_later = 2)"
		objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10,var_cart_userid))
		Set rs_getAnodizedItems = objCmd.Execute()
			var_anodizationPrice = var_anodizationBasePrice
	End If	
End If

var_anodizationSubTotal = var_anodizationPrice * rs_getCart.Fields.Item("cart_qty").Value
var_lineTotal = var_itemPrice * rs_getCart.Fields.Item("cart_qty").Value 
var_lineTotal = var_lineTotal + var_anodizationSubTotal
var_lineTotal_subtotal = var_itemPrice_USdollars * rs_getCart.Fields.Item("cart_qty").Value
var_subtotal = var_subtotal + var_lineTotal_subtotal + var_anodizationSubTotal

var_couponTotal = var_couponTotal + var_line_coupon
total_preferred_discount = FormatNumber(total_preferred_discount + var_discount_preferred, -1, -2, -2, -2)	
%>