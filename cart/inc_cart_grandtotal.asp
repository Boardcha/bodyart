<% 
' used by the IPGeo program to estimate shipping cost before they select a method at checkout
if strcountryName <> "" then 
		shipping_cost = var_shipping_cost ' variable set on /cart/inc_cart_loopitems-end.asp
end if

' Only used for /checkout/ajax_process_payment.asp for the *exact* price paid on shipping
if session("shipping_cost") <> "" then 
	shipping_cost = session("shipping_cost") ' variable set on /checkout/inc_store_shipping_selection.asp
	session("temp_shipping") = 0
end if


if session("credit_now") = "" then
	var_credit_now = 0
else
	var_credit_now = FormatNumber(session("credit_now"),2)
end if
if session("credit_later") = "" then
	session("credit_later") = 0
end if

' Set variables
var_total_without_certsOrCredits = var_subtotal - var_couponTotal - total_preferred_discount - var_credit_now + var_salesTax + shipping_cost

var_credit_due_todb = 0



' ==============================================
' START do some prioritization if certs AND store credits are trying to be used at the same time
if Session("GiftCertAmount") <> 0 and session("usecredit") <> "" then



	' LOGIC USED 
	'if gift cert is <  the total
	'	use up the gift cert
	'	store credit, write balance due, or $0 to db


	
	if Session("GiftCertAmount") < var_total_without_certsOrCredits then
		var_total_giftcert_dueback = 0
		var_total_giftcert_used = Session("GiftCertAmount")
		
		'Write remaining balance unused for store credit
		session("storeCredit_used") = FormatNumber(var_total_without_certsOrCredits, -1, -2, -2, -2) - FormatNumber(var_total_giftcert_used, -1, -2, -2, -2)
	
		
		' set it to where store credit can't exceed what's on their account
		if session("storeCredit_used") > session("storeCredit_amount") then
			session("storeCredit_used") = FormatNumber(session("storeCredit_amount"), -1, -2, -2, -2)
		end if
		
		' Variable for amount due back to write to database
		var_credit_due_todb = FormatNumber(session("storeCredit_amount") - session("storeCredit_used"), -1, -2, -2, -2)	
	
	
	end if ' END if gift cert is <  the total

	

	'if gift cert is >= the total
	'	write gift cert remaining balance due
	'	don't touch the store credit

	
	if Session("GiftCertAmount") >= var_total_without_certsOrCredits then
		
		'Write remaining balance unused for gift certs
			var_total_giftcert_used = var_total_without_certsOrCredits
		
	'	var_shipping_cost_friendly = var_total_giftcert_used ' BUG TESTING
	
			' set it to where gift cert can't exceed what's on their account
			if var_total_giftcert_used > session("GiftCertAmount") then
				var_total_giftcert_used = FormatNumber(session("GiftCertAmount"), -1, -2, -2, -2)
			end if
			
			' Variable for amount due back to write to database
			var_total_giftcert_dueback = FormatNumber(session("GiftCertAmount") - var_total_giftcert_used, -1, -2, -2, -2)
			
	'	session("storeCredit_amount") = 0
		session("storeCredit_used") = 0
		var_credit_due_todb = session("storeCredit_amount")
		
	end if ' END if gift cert is >= the total
	
end if
' ==============================================
' END do some prioritization if certs AND store credits are trying to be used at the same time

' -----------------------------------------------
' START If JUST a gift cert is being used AND NOT with a store credit
if Session("GiftCertAmount") <> 0 and session("usecredit") = "" then

			'Write remaining balance unused for gift certificate
			if session("GiftCertAmount") >= var_total_without_certsOrCredits then
				var_total_giftcert_used = var_total_without_certsOrCredits
			else
				var_total_giftcert_used = FormatNumber(session("GiftCertAmount"), -1, -2, -2, -2)
			end if
							
			' Variable for amount due back to write to database
			var_total_giftcert_dueback = FormatNumber(session("GiftCertAmount") - var_total_giftcert_used, -1, -2, -2, -2)
			

end if  ' ----------------------------------------
' END If JUST a store credit is being used AND NOT with a gift cert

' -----------------------------------------------
' START If JUST a STORE CREDIT is being used AND NOT with a gift cert
if Session("GiftCertAmount") = 0 and session("usecredit") <> "" then

'response.write "<br/>var_total_without_certsOrCredits " & var_total_without_certsOrCredits
'response.write "<br/>credit amount " & FormatNumber(session("storeCredit_amount"))
'response.write "<br/>used " & session("storeCredit_used")

			'Write remaining balance unused for store credit
			if session("storeCredit_amount") >= var_total_without_certsOrCredits then
				session("storeCredit_used") = var_total_without_certsOrCredits
			else
				session("storeCredit_used") = FormatNumber(session("storeCredit_amount"), -1, -2, -2, -2)
			end if
							
			' Variable for amount due back to write to database
			var_credit_due_todb = FormatNumber(session("storeCredit_amount") - session("storeCredit_used"), -1, -2, -2, -2)	
			

end if  ' ----------------------------------------
' END If JUST a gift certificate is being used AND NOT with a store credit

var_grandtotal = ((var_subtotal - var_couponTotal - total_preferred_discount - var_credit_now + var_salesTax + shipping_cost - session("storeCredit_used")) - var_total_giftcert_used)

if session("amount_to_collect") <> "" then
if session("amount_to_collect") <> 0 then
	var_grandtotal = ((var_subtotal - var_couponTotal - total_preferred_discount - var_credit_now + session("amount_to_collect") + shipping_cost - session("storeCredit_used")) - var_total_giftcert_used)
end if
end if

	
if var_grandtotal < 0 then
	var_grandtotal = 0
end if

if strcountryName <> "" then 
		session("temp_shipping") = ((var_subtotal - var_couponTotal - total_preferred_discount - var_credit_now + var_salesTax + shipping_cost - session("storeCredit_used")) - var_total_giftcert_used) ' used on checkout.asp page to hide payment section on page load faster
end if
%>