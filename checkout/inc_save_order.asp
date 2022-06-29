<!--#include virtual="/functions/function-decode-to-utf8.asp" -->
<%
if var_addons_active <> "yes" then

	' START if customer is logged in then get CIM auth.net address information
	' =================================================================================
	if CustID_Cookie <> "" and CustID_Cookie <> 0 then

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM customers WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
		Set rsGetUser = objCmd.Execute()
		
		
		' Set customers main CIM ID to variable
		var_cim_custid = rsGetUser.Fields.Item("cim_custid").Value
		var_email = rsGetUser.Fields.Item("email").Value
		var_our_custid = rsGetUser.Fields.Item("customer_ID").Value

			' used to store into order on checkout/inc_save_cims_to_order
			var_cim_shipping_id = request.form("cim_shipping")
		
		' Get CIM SHIPPING information ONLY if a new address wasn't entered into form
		if request.form("cim_shipping") <> "" and (request.form("shipping-first") = "" and request.form("shipping-last") = "") then
		
			' Connect to Authorize.net CIM to get shipping address book information
			strGetShipping = "<?xml version=""1.0"" encoding=""utf-8""?>" _
			& "<getCustomerShippingAddressRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
			& MerchantAuthentication() _
			& "  <customerProfileId>" & var_cim_custid & "</customerProfileId>" _
			& "  <customerAddressId>" & request.form("cim_shipping") & "</customerAddressId>" _
			& "</getCustomerShippingAddressRequest>"
			
			Set objResponseGetShipping = SendApiRequest(strGetShipping)

			' If connection is a success to address book than retrieve values and assign to variables
			If IsApiResponseSuccess(objResponseGetShipping) Then
				var_shipping_first = DecodeUTF8(objResponseGetShipping.selectSingleNode("/*/api:address/api:firstName").Text)
				
				session("shipping_first") = DecodeUTF8(objResponseGetShipping.selectSingleNode("/*/api:address/api:firstName").Text)
		
				var_shipping_last = DecodeUTF8(objResponseGetShipping.selectSingleNode("/*/api:address/api:lastName").Text)
				
				session("shipping_last") = DecodeUTF8(objResponseGetShipping.selectSingleNode("/*/api:address/api:lastName").Text)
				
				var_shipping_company = DecodeUTF8(objResponseGetShipping.selectSingleNode("/*/api:address/api:company").Text)
				
				session("shipping_company") = DecodeUTF8(objResponseGetShipping.selectSingleNode("/*/api:address/api:company").Text)
				
				'Split out state from authorize.net and break it out to address 1 and address 2 fields
				split_address = Split(objResponseGetShipping.selectSingleNode("/*/api:address/api:address").Text, "|")
					var_shipping_address1 = DecodeUTF8(split_address(0))
					var_shipping_address2 = DecodeUTF8(split_address(1))
					session("shipping_address1") = DecodeUTF8(split_address(0))
					session("shipping_address2") = DecodeUTF8(split_address(1))
					
					
						
				var_shipping_city = DecodeUTF8(objResponseGetShipping.selectSingleNode("/*/api:address/api:city").Text)
				
				session("city") = DecodeUTF8(objResponseGetShipping.selectSingleNode("/*/api:address/api:city").Text)
				
				var_shipping_state = replace(DecodeUTF8(objResponseGetShipping.selectSingleNode("/*/api:address/api:state").Text), "|", "")
				
				session("state") =  replace(DecodeUTF8(objResponseGetShipping.selectSingleNode("/*/api:address/api:state").Text), "|", "")
					
				var_shipping_zip = objResponseGetShipping.selectSingleNode("/*/api:address/api:zip").Text
				
				session("shipping_zip") = objResponseGetShipping.selectSingleNode("/*/api:address/api:zip").Text
				
				var_shipping_country = objResponseGetShipping.selectSingleNode("/*/api:address/api:country").Text
				
				session("country") = objResponseGetShipping.selectSingleNode("/*/api:address/api:country").Text
				
				var_phone = objResponseGetShipping.selectSingleNode("/*/api:address/api:phoneNumber").Text
				
				strShipping_ID = objResponseGetShipping.selectSingleNode("/*/api:address/api:customerAddressId").Text
			End if
		
		end if ' Get CIM shipping information ONLY if a new address wasn't entered into form
		
		' START Get CIM BILLING information ONLY if a new address wasn't entered into form
		if request.form("cim_billing") <> "" and (request.form("billing-first") = "" and request.form("billing-last") = "") then
		
			' used to store into order on checkout/inc_save_cims_to_order
			var_cim_billing_id = request.form("cim_billing")
		
			strGetBillingAddress = "<?xml version=""1.0"" encoding=""utf-8""?>" _
			& "<getCustomerPaymentProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
			& MerchantAuthentication() _
			& "  <customerProfileId>" & cim_custid & "</customerProfileId>" _
			& "  <customerPaymentProfileId>" & request.form("cim_billing") & "</customerPaymentProfileId>" _
			& "</getCustomerPaymentProfileRequest>"
			
			Set objResponseGetAddress = SendApiRequest(strGetBillingAddress)

			' If connection is a success to address book than retrieve values and assign to variables
			If IsApiResponseSuccess(objResponseGetAddress) Then
				strBilling_cardnumber = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:payment/api:creditCard/api:cardNumber").Text
				strCardType = "Credit card"			
		' NOT VALID BY AUTH.NET	strCardType = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:payment/api:creditCard/api:cardType").Text
				'Split out state from authorize.net and break it out to address 1 and address 2 fields
				split_address_billing = Split(objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:address").Text, "|")
					var_billing_address1 = split_address_billing(0)
					var_billing_address = split_address_billing(0)
					var_billing_address2 = split_address_billing(1)
						
				var_billing_city = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:city").Text
				var_billing_state = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:state").Text
				var_billing_zip = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:zip").Text
				strBilling_ID = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:customerPaymentProfileId").Text
			End if
			
		end if ' END Get CIM BILLING information ONLY if a new address wasn't entered into form

	else ' IF CUSTOMER IS NOT LOGGED INSERT

		var_email = request.form("e-mail")
		
	end if
	' END getting CIM information
	' =================================================================================
	session("email") = var_email

	' Get form values IF customer is not logged in or information has been entered into form
	if request.form("shipping-first") <> "" and request.form("shipping-last") <> "" then
		
		var_phone = request.form("phone")
		var_shipping_company = request.form("shipping-company")
		var_shipping_first = request.form("shipping-first")
		var_shipping_last = request.form("shipping-last")
		var_shipping_address1 = request.form("shipping-address")
		var_shipping_address2 = request.form("shipping-address2")
		var_shipping_city = request.form("shipping-city")
		session("city") =  request.form("shipping-city")
		var_shipping_state = request.form("shipping-state")
		session("state") = request.form("shipping-state")
		
		var_shipping_province = request.form("shipping-province-canada") & "" & request.form("shipping-province")
		var_shipping_zip = request.form("shipping-zip")
		var_shipping_country = request.form("shipping-country")
		session("country") = request.form("shipping-country")
		var_shipping_phone = request.form("shipping-phone")
		
		' variables that need to be created to send out an email (specifically for PayPal.. needs to be session saved)
		
		session("shipping_company") = request.form("shipping-company")
		session("shipping_first") = request.form("shipping-first")
		session("shipping_last") = request.form("shipping-last")
		session("shipping_address1") = request.form("shipping-address")
		session("shipping_address2") = request.form("shipping-address2")
		session("shipping_province") = request.form("shipping-province-canada") & "" & request.form("shipping-province")
		session("shipping_zip") = request.form("shipping-zip")

	end if	

	if request.form("billing-first") <> "" and request.form("billing-last") <> "" then

		var_billing_name = request.form("billing-first") & " " & request.form("billing-last")
		var_billing_address = request.form("billing-address")
		var_billing_zip = request.form("billing-zip")

	end if
	
	'Detect card # text and assing it proper card type
	if request.form("card_number") <> "" then
		if Left(request.form("card_number"), 1) = 3 then ' AMEX
			strCardType = "American Express"
		end if
		if Left(request.form("card_number"), 1) = 4 then ' Visa
			strCardType = "Visa"
		end if
		if Left(request.form("card_number"), 1) = 5 then ' Mastercard
			strCardType = "Mastercard"
		end if
		if Left(request.form("card_number"), 1) = 6011 then ' Discover
			strCardType = "Discover"
		end if
	end if
	
	If request.form("googlepay") = "on" Or request.form("applepay") = "on" Then 'In this case, shipping and billing info comes from Google and Apple API
	    var_billing_name = request.form("full_name")
		var_shipping_first = getFirstName(request.form("full_name"))
		var_shipping_last = getLastName(request.form("full_name"))
		var_billing_address = request.form("address1")
		var_billing_zip = request.form("postal_code")
		var_shipping_address1 = request.form("address1")
		var_shipping_address2 = request.form("address2")
		var_shipping_city = request.form("locality")
		var_shipping_state = request.form("administrative_area")
		var_shipping_zip = request.form("postal_code")	
		var_shipping_phone = request.form("phone_number")		
		var_shipping_country_code = request.form("country_code")
		var_email = request.form("email")
		session("email") = var_email
		
		session("shipping_first") = var_shipping_first
		session("shipping_last") = var_shipping_last
		session("shipping_address1") = var_shipping_address1
		session("shipping_address2") = var_shipping_address2
		session("shipping_province") = request.form("shipping-province-canada") & "" & request.form("shipping-province")
		session("shipping_zip") = var_shipping_zip	
		session("city") =  var_shipping_city
		session("state") = var_shipping_state
		
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT Country, Country_UPSCode FROM TBL_Countries WHERE Country_UPSCode = '" & var_shipping_country_code & "'"
		Set rsCountry = objCmd.Execute()

		If Not rsCountry.EOF	Then
			var_shipping_country = rsCountry("Country")
		Else
			var_shipping_country = var_shipping_country_code
			
		End If
		session("country") = var_shipping_country
		Set rsCountry = Nothing
	End If
	
end if ' if var_addons_active 

' CREDIT CARD payment method
' =================================================================================
if request.form("card_number") <> "" or (request.form("cim_billing") <> "paypal" and request.form("cim_billing") <> "cash" and request.form("afterpay") <> "on" and request.form("googlepay") <> "on" and request.form("applepay") <> "on") then
%>
	"paypal":"no",
	"afterpay":"no",
	"googlepay":"no",
	"applepay":"no",
	"cash":"no",
	"invoiceid":"0",
	"order":"saved",
	"status":"",
<%  
end if ' credit card payment
' =================================================================================


' PAYPAL payment method
' =================================================================================
	if request.form("paypal") = "on" or request.form("cim_billing") = "paypal" then
	
		strCardType = "PayPal"
	%>
		"paypal":"yes",
		"afterpay":"no",
		"googlepay":"no",
		"applepay":"no",
		"cash":"no",
		"order":"saved",
		"status":"",
		"cc_approved":"no"		
	<%	
	end if ' if payment method is paypal
' =================================================================================

' AFTERPAY payment method
' =================================================================================
	if request.form("afterpay") = "on" then
	
		strCardType = "Afterpay"
	%>
		"paypal":"no",
		"afterpay":"yes",
		"googlepay":"no",
		"applepay":"no",		
		"cash":"no",
		"order":"saved",
		"status":"",
		"cc_approved":"no"		
	<%	
	end if ' if payment method is afterpay
' =================================================================================

' CASH payment method
' =================================================================================
	if request.form("cash") = "on" or request.form("cim_billing") = "cash" then
	
		strCardType = "Cash"
	%>
		"cash":"yes",
		"paypal":"no",
		"googlepay":"no",
		"applepay":"no",		
		"order":"saved",
		"status":"cash",
		"cc_approved":"no"
	<%	
	end if ' if payment method is CASH
' =================================================================================

' Google Pay payment method
' =================================================================================
	if request.form("googlepay") = "on" then
	
		strCardType = "GooglePay"
	%>
		"cash":"no",
		"paypal":"no",
		"googlepay":"yes",
		"applepay":"no",		
		"order":"saved",
		"status":"",
	<%	
	end if ' if payment method is GooglePay
' =================================================================================

' Apple Pay payment method
' =================================================================================
	if request.form("applepay") = "on" then
	
		strCardType = "ApplePay"
	%>
		"cash":"no",
		"paypal":"no",
		"googlepay":"no",
		"applepay":"yes",		
		"order":"saved",
		"status":"",
	<%	
	end if ' if payment method is ApplePay
' =================================================================================

if var_addons_active <> "yes" then

' BEGIN STORE ORDER -- regardless of payment method
' =================================================================================


	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	'objCmd.CommandType = 4
	'objCmd.CommandText = "Proc_Checkout4_InsertOrder"
	objCmd.CommandText = "INSERT INTO sent_items (customer_ID, email, company, customer_first, customer_last, address, address2, city, state, province, zip, country, phone, UPS_AmountPaid, UPS_Service, item_description, customer_comments, shipping_rate, shipping_type, pay_method, shipped, date_order_placed, coupon_code, IPaddress,  billing_name, billing_address, billing_zip, preorder, autoclave, total_sales_tax, taxes_state_only, taxes_county_only, taxes_city_only, taxes_special_only, combined_tax_rate, total_gift_cert, total_coupon_discount, total_preferred_discount, total_store_credit, total_free_credits, giftcert_flag, currency_type, exchange_rate, checkout_estimated_delivery_date, anodize, referral_traffic_from, referral_traffic_to) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

	var_customer_comments = replace(request.form("Comments"), ",", "")
	var_gift_order = replace(request.form("gift"), ",", "")
	var_conserve = replace(request.form("conserve-plastic"), ",", "")
	
	' Build customer comments to be stored in public notes field
	if request.form("Comments") <> "" then
		session("customer_comments") = var_customer_comments
		strCust_Comments = "<br/>CUSTOMER COMMENTS:<br/>" & Replace(var_customer_comments, vbCrLF, "<br />" + vbCrLF) & "<br/><br/>"
	end if
	
	
	objCmd.NamedParameters = True	
	objCmd.Parameters.Append(objCmd.CreateParameter("@CustomerID",3,1,10,var_our_custid)) ' set on this page and on inc_create_account.asp depending on if they are logged in or create a new account
	objCmd.Parameters.Append(objCmd.CreateParameter("@Email",200,1,70,var_email))
			
			
			' MOVE TO SAVE ORDER			objCmd.Parameters.Append(objCmd.CreateParameter("@cim_id", 200,1,30, session("cim_accountNumber")))
' MOVE TO SAVE ORDER						objCmd.Parameters.Append(objCmd.CreateParameter("@shipping_profile_id", 200,1,30, Session("shipping_profile_id")))
' MOVE TO SAVE ORDER						objCmd.Parameters.Append(objCmd.CreateParameter("@payment_profile_id", 200,1,30, Session("billing_cim_address_id")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@Company",200,1,50,var_shipping_company))
			objCmd.Parameters.Append(objCmd.CreateParameter("@First",200,1,30, replace(var_shipping_first,"""", "")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@Last",200,1,30, replace(var_shipping_last,"""", "")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@Street",200,1,75,replace(var_shipping_address1,"""", "")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@Address2",200,1,75,replace(var_shipping_address2,"""", "")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@City",200,1,50, replace(var_shipping_city,"""", "")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@State",200,1,50, replace(var_shipping_state,"""", "")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@Province",200,1,30,replace(var_shipping_province,"""", "")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@Zip",200,1,15, replace(var_shipping_zip,"""", "")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@Country",200,1,100,var_shipping_country))
			objCmd.Parameters.Append(objCmd.CreateParameter("@Phone",200,1,20,var_shipping_phone))		
			objCmd.Parameters.Append(objCmd.CreateParameter("@UPS_AmountPaid",6,1,10,ups_shipping_actual))
			objCmd.Parameters.Append(objCmd.CreateParameter("@UPS_Service",200,1,30,ups_shipping_type))
			objCmd.Parameters.Append(objCmd.CreateParameter("@item_description",200,1,1000,var_gift_order+""+var_conserve + strCust_Comments))
			objCmd.Parameters.Append(objCmd.CreateParameter("@customer_comments",200,1,2000,Replace(var_customer_comments, vbCrLF, "<br />" + vbCrLF)))		
			objCmd.Parameters.Append(objCmd.CreateParameter("@shipping_rate",6,1,10,session("shipping_cost")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@shipping_type",200,1,70,shipping_option))
			
			objCmd.Parameters.Append(objCmd.CreateParameter("@pay_method",200,1,30,strCardType))
			objCmd.Parameters.Append(objCmd.CreateParameter("@shipped",200,1,50,"Pending..."))

			'objCmd.Parameters.Append(objCmd.CreateParameter("@date_order_placed",200,1,30,now())) 'UGUR: This doesn't work on my local, see below line
			objCmd.Parameters.Append(objCmd.CreateParameter("@date_order_placed",200,1,30,Cstr(now())))

			if session("preferred") = "yes" then
				var_store_coupon = "YTG89R57"
			end if
			if Session("CouponCode") <> "" then
				var_store_coupon = Session("CouponCode")
			end if
			
			objCmd.Parameters.Append(objCmd.CreateParameter("@coupon_code",200,1,30,var_store_coupon))
			objCmd.Parameters.Append(objCmd.CreateParameter("@IPaddress",200,1,30,Request.ServerVariables("REMOTE_HOST")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@billing_name",200,1,30,var_billing_name))
			objCmd.Parameters.Append(objCmd.CreateParameter("@billing_address",200,1,75,var_billing_address))
			objCmd.Parameters.Append(objCmd.CreateParameter("@billing_zip",200,1,15,var_billing_zip))

			if preorder_shipping_notice = "yes" then
				objCmd.Parameters.Append(objCmd.CreateParameter("@preorder",3,1,1,1))
			else
				objCmd.Parameters.Append(objCmd.CreateParameter("@preorder",3,1,1,0))
			end if
			
			if var_sterilization_added = 1 then
				objCmd.Parameters.Append(objCmd.CreateParameter("@autoclave",3,1,1,1))
			else
				objCmd.Parameters.Append(objCmd.CreateParameter("@autoclave",3,1,1,0))
			end if	
			
			objCmd.Parameters.Append(objCmd.CreateParameter("@total_sales_tax",6,1,10,session("amount_to_collect")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@taxes_state_only",6,1,10,session("state_tax_collectable")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@taxes_county_only",6,1,10,session("county_tax_collectable")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@taxes_city_only",6,1,10,session("city_tax_collectable")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@taxes_special_only",6,1,10,session("special_district_tax_collectable")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@combined_tax_rate",6,1,10,session("combined_tax_rate")))
			objCmd.Parameters.Append(objCmd.CreateParameter("@total_gift_cert",6,1,10,FormatNumber(var_total_giftcert_used, -1, -2, -2, -2)))
			objCmd.Parameters.Append(objCmd.CreateParameter("@total_coupon_discount",6,1,10,FormatNumber(var_couponTotal, -1, -2, -2, -2)))
			objCmd.Parameters.Append(objCmd.CreateParameter("@total_preferred_discount",6,1,10,FormatNumber(total_preferred_discount, -1, -2, -2, -2)))
			objCmd.Parameters.Append(objCmd.CreateParameter("@total_store_credit",6,1,10,FormatNumber(session("storeCredit_used"),2)))
			objCmd.Parameters.Append(objCmd.CreateParameter("@total_free_credits",6,1,10,FormatNumber(var_credit_now,2)))
			
			if var_giftcert = "yes" then
				objCmd.Parameters.Append(objCmd.CreateParameter("@giftcert_flag",3,1,1,1))
			else
				objCmd.Parameters.Append(objCmd.CreateParameter("@giftcert_flag",3,1,1,0))
			end if

			objCmd.Parameters.Append(objCmd.CreateParameter("@currency_type",200,1,20, session("exchange-currency") ))

			if session("exchange-rate") <> "" then
				objCmd.Parameters.Append(objCmd.CreateParameter("@exchange_rate",6,1,10, session("exchange-rate") ))
			else
				objCmd.Parameters.Append(objCmd.CreateParameter("@exchange_rate",6,1,10, 0 ))
			end if
			
			if session("EXP_checkout_estimated_delivery") <> "" Then
				objCmd.Parameters.Append(objCmd.CreateParameter("@checkout_estimated_delivery",200,1,30, session("EXP_checkout_estimated_delivery")))
			elseif session("MAX_checkout_estimated_delivery") <> "" then
				objCmd.Parameters.Append(objCmd.CreateParameter("@checkout_estimated_delivery",200,1,30, session("MAX_checkout_estimated_delivery")))
			else	
				objCmd.Parameters.Append(objCmd.CreateParameter("@checkout_estimated_delivery",200,1,30, NULL ))
			end if			
			
			if var_anodization_added = 1 then
				objCmd.Parameters.Append(objCmd.CreateParameter("@anodize",3,1,1,1))
			else
				objCmd.Parameters.Append(objCmd.CreateParameter("@anodize",3,1,1,0))
			end if	
			
			if Session("referral_traffic_from") <> "" Then
				objCmd.Parameters.Append(objCmd.CreateParameter("@referral_traffic_from",200,1,200, Session("referral_traffic_from")))
			else	
				objCmd.Parameters.Append(objCmd.CreateParameter("@referral_traffic_from",200,1,200, NULL ))
			end if		

			if Session("referral_traffic_to") <> "" Then
				objCmd.Parameters.Append(objCmd.CreateParameter("@referral_traffic_to",200,1,200, Session("referral_traffic_to")))
			else	
				objCmd.Parameters.Append(objCmd.CreateParameter("@referral_traffic_to",200,1,200, NULL ))
			end if				
			
	objCmd.Execute()

	' Get invoice # for order
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP 5 ID FROM sent_items WHERE email = ? ORDER BY ID DESC"
	objCmd.Parameters.Append(objCmd.CreateParameter("Email",200,1,70,var_email))
	Set rsGetInvoiceNum = objCmd.Execute()

		'STORE INVOICE # TO SESSION ----------------------------------
		Session("invoiceid") = rsGetInvoiceNum.Fields.Item("ID").Value
	
	set rsGetInvoiceNum = nothing
	
	While Not rs_getCart.EOF
		
		var_preorder_text = ""
		var_itemPrice = 0
		'If ProductID flagged as "waiting-list", meaning if customer comes from waiting-list email notification, save this info to the "referrer" field.
		If Session(rs_getCart("ProductID")) = "waiting-list" Then 
			var_referrer = "'waiting-list'" 
		ElseIf rs_getCart("cart_save_for_later") = 2 Then ' 2 = it is added to cart back from saved items
			var_referrer = "'save-for-later'" 
		Else 
			var_referrer = "NULL"
		End If

		var_SavePrice_USdollars =  FormatNumber(rs_getCart.Fields.Item("price").Value, -1, -2, -2, -2)
		if (rs_getCart("SaleDiscount") > 0 AND rs_getCart("secret_sale") = 0) OR (rs_getCart("secret_sale") = 1 AND session("secret_sale") = "yes") then
			var_SavePrice_USdollars = FormatNumber(var_SavePrice_USdollars - (rs_getCart("price") * rs_getCart("SaleDiscount")/100), -1, -2, -2, -2)
		end if

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM TBL_Anodization_Colors_Pricing WHERE anodID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("id",200,1,70,rs_getCart.Fields.Item("anodID").Value))
		Set rsAnodizeFee = objCmd.Execute()

		If Not rsAnodizeFee.EOF Then 
			var_anodization_fee = rsAnodizeFee("base_price")
		else
			var_anodization_fee = 0
		end if


		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, ProductID, DetailID, qty, item_price,  PreOrder_Desc, anodization_id_ordered, item_wlsl_price, referrer, anodization_fee) VALUES (?,?,?,?,?,?,?,?," & var_referrer & ", ?)"
				objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,session("invoiceid")))
				objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15, rs_getCart("ProductID") ))
				objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15, rs_getCart("ProductDetailID") ))
				objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10, rs_getCart("cart_qty") ))
				objCmd.Parameters.Append(objCmd.CreateParameter("item_price",6,1,10, FormatNumber(var_SavePrice_USdollars, -1, -2, -2, -2) ))
				objCmd.Parameters.Append(objCmd.CreateParameter("preorder_notes",200,1,2000, replace(rs_getCart("cart_preorderNotes"),"{}", "   ") ))
				objCmd.Parameters.Append(objCmd.CreateParameter("anodization_id_ordered",3,1,15, rs_getCart("anodID") ))
				objCmd.Parameters.Append(objCmd.CreateParameter("item_wlsl_price",6,1,10, FormatNumber(rs_getCart("wlsl_price"), -1, -2, -2, -2) ))
				objCmd.Parameters.Append(objCmd.CreateParameter("anodization_fee",6,1,10, var_anodization_fee ))
		objCmd.Execute()
	rs_getCart.MoveNext()
	wend
	
' END store order
' =================================================================================
end if 'if var_addons_active
%>

<%
' FUNCTIONS
Function getFirstName(fullName)

	If Instr(fullName, ",") > 0 Then
		firstName = Trim(Mid(fullName, Instr(fullName, ",") + 1))
	ElseIf Instr(fullName, " ") > 0 Then
		firstName = Trim(Mid(fullName, 1, InstrRev(fullName, " ")))
	Else
		firstName = fullName
	End If
	
	getFirstName = firstName
	
End Function

Function getLastName(fullName)

	If Instr(fullName, ",") > 0 Then
		lastName = Trim(Mid(fullName, 1, Instr(fullName, ",") - 1))
	ElseIf Instr(fullName, " ") > 0 Then
		lastName = Trim(Mid(fullName, InstrRev(fullName, " ") + 1))
	Else
		lastName = ""
	End If
	
	getLastName = lastName
	
End Function
%>
