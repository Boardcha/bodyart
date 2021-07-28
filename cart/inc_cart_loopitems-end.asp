<%		
rs_getCart.MoveNext()
Loop
' End recordset list

' Auto remove autoclave service if nothing is in cart that can be autoclaved
if var_autoclavable = 0 and var_cart_id_autoclave <> "" then
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "DELETE FROM tbl_carts WHERE cart_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("cart_id",3,1,10,var_cart_id_autoclave))
	objCmd.Execute()
	

end if

' Set variable that is used on cart/fraud/inc_freegifts_doublechecks.asp
fraudcheck_freegifts_subtotal = var_subtotal - var_couponTotal - total_preferred_discount
%>
<!--#include file="fraud_checks/inc_freegifts_doublechecks.asp"-->
<%

'--------- add up use now credits for free gifts if selected
	' DO NOT RUN ON PROCESS ORDER PAGE SINCE SESSION VALUES SHOULD ALREADY BE SET ... causes timeouts and errors
	credit_later = 0
if var_process_order <> "yes" then
	for loop_free_cookie = 1 to 5
		' display free gifts file 5 times
	' retrieve and total use now free credits			
	Do While Not rsGetFree.EOF

		if cStr(rsGetFree.Fields.Item("ProductDetailID").Value) = request.cookies("freegift" & loop_free_cookie & "id") then
		
			'Detect whether to use a credit now, or save for later (if customer selected that option)
			if Instr(1, lcase(rsGetFree.Fields.Item("ProductDetail1")), lcase("USE NOW")) > 0 Then
				credit_now = credit_now + rsGetFree.Fields.Item("price")		
			end if
			
			if Instr(1, lcase(rsGetFree.Fields.Item("ProductDetail1")), lcase("LATER")) > 0 Then
				credit_later = credit_later + rsGetFree.Fields.Item("price")
			' response.write credit_later
			end if

		end if ' find matching information for stored cookie id
	  rsGetFree.MoveNext()
	Loop
		rsGetFree.MoveFirst()
	next  'for loop_free_cookie = 1 to 5
		
		session("credit_now") = credit_now
		session("credit_later") = credit_later
		
		
		
end if 	' if var_process_order <> "yes" ---  DO NOT RUN ON PROCESS ORDER PAGE SINCE SESSION VALUES SHOULD ALREADY BE SET ... causes timeouts and errors

'--------- add up use now credits for free gifts if selected

' calculate footer items
'var_subtotal_after_discounts --
' two certs (one for buying them so they don't get charged tax, and then one for using them them to reduce the taxed amount)
var_subtotal_after_discounts = FormatNumber(var_subtotal  - var_totalvalue_certs_incart - var_couponTotal - total_preferred_discount - session("credit_now"), -1, -2, -2, -2)
var_total_discounts = FormatNumber(var_couponTotal + Session("GiftCertAmount") + session("storeCredit_amount"), -1, -2, -2, -2)
var_total_discounts_noStoreCredit =  FormatNumber(var_couponTotal + Session("GiftCertAmount"), -1, -2, -2, -2)

session("taxable_amount") = var_subtotal_after_discounts
%>
<!--#include virtual="cart/inc_display_estimated_shipping.asp"-->
<%
if session("shipping-state") = "" then
	if request.form("state") <> "" then
		session("shipping-state") = request.form("state")
	else
		if request.cookies("ip-region") <> "" then
			session("shipping-state") = request.cookies("ip-region")
		else
			session("shipping-state") = ""
		end if
	end if
end if

if session("amount_to_collect") <> "" then	
	if session("amount_to_collect") = 0 then
		var_salesTax = 0
	else
		var_salesTax = session("amount_to_collect")
	end if
else
	var_salesTax = 0
end if



'response.write var_totalvalue_certs_incart
'response.write ", " & var_subtotal_after_discounts

'response.write var_totalvalue_certs_incart
' have to subtract Session("GiftCertAmount") * 2 so that it doesn't count AGAINST the customer
var_shipping_AmountNeeded = FormatNumber((25 - var_subtotal_after_discounts), -1, -2, -2, -2)
' do not show 

'response.write var_shipping_AmountNeeded & " , " & var_totalvalue_certs_incart


				
' Reset shipping if there's ONLY a GIFT CERTIFICATE in the cart			
if var_other_items <> 1 then 
	var_only_gift_cert = "yes"
	var_shipping_cost = 0
	shipping_cost = 0
	session("shipping_cost") = 0
	var_shipping_cost_friendly ="FREE"
	var_salesTax = 0
end if

' Reset shipping if add-on items are being added to an order that has not shipped out yet
if request.cookies("OrderAddonsActive") <> "" then
	var_shipping_cost = 0
	shipping_cost = 0
	session("shipping_cost") = 0
	var_shipping_cost_friendly = "Paid on original order"
	session("var_email_shipping_option") = "Paid on original order"
end if

'=============================================================================================

%>
