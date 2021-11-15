<%
' CREATE NEW CUSTOMER ACCOUNT if they opted in
' WHEN EDITING THIS PAGE MAKE SURE TO TEST CHECKOUT NEW ACCOUNT CREATION AS WELL AND MAKE SURE IT DOES NOT BREAK CHECKOUT
' =================================================================================
password = request.form("password")
check = request.form("check")
email = request.form("e-mail")
'overwrite variables if user signed-in with Google
If google_signin_email <> "" Then
	email = google_signin_email
	password = "hg4!=g4s68.n" & getSalt(10, extraChars) 'Set a different temporary password for each user signed in with Google
End If

'====== CHECKING FOR PASSWORD MUST BE LEFT HERE OTHERWISE IT'LL BREAK CHECKOUT
if password <> "" and email <> "" and check = "" then
	' THe check field is to help prevent against bots

	salt = getSalt(32, extraChars)
	newPass = sha256(salt & password & extra_key)
	activation_hash = getToken(32, "")
	'Add new account information into our database
	If google_firstName <> "" Then firstName = google_firstName Else firstName = NULL
	If google_lastName <> "" Then lastName = google_lastName Else lastName = NULL
	If google_user_id <> "" Then 
		googleUserId = google_user_id
		activate_account = 1
		'Set variable for mailer
			mailer_type = "new account"
		' Set extra mailer type
			email_onetime_coupon = "yes"
%>
			<!--#include virtual="/checkout/inc_random_code_generator.asp"-->
			<!--#include virtual="/includes/inc-dupe-onetime-codes.asp"--> 
<%

			' Prepare a one time use coupon for creating an account
			var_cert_code = getPassword(15, extraChars, firstNumber, firstLower, firstUpper, firstOther, latterNumber, latterLower, latterUpper, latterOther)

			' Call function
			var_cert_code = CheckDupe(var_cert_code)

			' Set extra mailer type
			email_onetime_coupon = "yes"

			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO TBLDiscounts (DiscountCode, DateExpired, coupon_single_email, DiscountPercent, coupon_single_use, DateAdded, DiscountType, active, dateactive, coupon_assigned, DiscountDescription) VALUES (?, GETDATE()+30, ?, 10, 1, GETDATE(), 'Percentage', 'A', GETDATE()-1, 1, 'New account creation')"
			objCmd.Parameters.Append(objCmd.CreateParameter("Code",200,1,30,var_cert_code))
			objCmd.Parameters.Append(objCmd.CreateParameter("Email",200,1,30, email))
			objCmd.Execute()
			'===== END IF GOOGLE USER IS CREATED


	Else googleUserId = NULL
		'Set variable for mailer
			mailer_type = "account activation"
		activate_account = 0
	End if
	If google_signin_email <> "" Then registered_with_social_login = 1 Else registered_with_social_login = 0
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO customers (email, password_hashed, salt, customer_first, customer_last, registered_with_social_login, google_user_id, activation_hash, active) VALUES (?, '" & newPass & "', '" & salt & "', ?, ?, ?, ?, '" & activation_hash & "', ?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,50, email))
	objCmd.Parameters.Append(objCmd.CreateParameter("firstName",200,1,50, firstName))
	objCmd.Parameters.Append(objCmd.CreateParameter("lastName",200,1,50, lastName))
	objCmd.Parameters.Append(objCmd.CreateParameter("registered_with_social_login",3,1,10,registered_with_social_login))
	objCmd.Parameters.Append(objCmd.CreateParameter("googleUserId",200,1,200, googleUserId))
	objCmd.Parameters.Append(objCmd.CreateParameter("activate_account",3,1,10,activate_account))
	objCmd.Execute()
	
	'Retrieve customer ID number from our database
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT customer_ID, email FROM customers WHERE email = ? ORDER BY customer_ID DESC"
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,50, email))
	Set rsGetUserID = objCmd.Execute()
	
	'Create cookie and session for new customer (log them in automatically)
		session("custID_account") = rsGetUserID.Fields.Item("customer_ID").Value
		var_our_custid = rsGetUserID.Fields.Item("customer_ID").Value
		
	'Write account creation and login dates
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET last_login = '" & now() & "', account_created = '" & now() & "' WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,rsGetUserID.Fields.Item("customer_ID").Value))
	objCmd.Execute()
	
	' Transfer any previous orders with matching email address AND custID = 0 over to new account
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET customer_ID = ? WHERE customer_ID = 0 AND email = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,rsGetUserID.Fields.Item("customer_ID").Value))
	objCmd.Parameters.Append(objCmd.CreateParameter("@Email",200,1,70,rsGetUserID.Fields.Item("email").Value))
	objCmd.Execute()
		
	' Authorize.net create customer profile
	Dim strReq
	Dim objResponse
	Dim strCustomerProfileId
	strCustomerProfileId = ""

	strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
	& "<createCustomerProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
	& MerchantAuthentication() _
	& "<profile>" _
	& "  <merchantCustomerId>" & rsGetUserID.Fields.Item("customer_ID").Value & "</merchantCustomerId>" _
	& "  <email>" & rsGetUserID.Fields.Item("email").Value & "</email>" _
	& "</profile>" _
	& "</createCustomerProfileRequest>"

	Set objResponseCreateProfile = SendApiRequest(strReq)

	' If succcess in created a new CIM profile ID then add that new ID to our database in the customers table
	If IsApiResponseSuccess(objResponseCreateProfile) Then
	  
	  strCustomerProfileId = objResponseCreateProfile.selectSingleNode("/*/api:customerProfileId").Text
	  
	  var_cim_custid = objResponseCreateProfile.selectSingleNode("/*/api:customerProfileId").Text
		
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "UPDATE customers SET cim_custid = ? WHERE customer_ID = ?"
			objCmd.Parameters.Append(objCmd.CreateParameter("cim_custid",200,1,30,Server.HTMLEncode(strCustomerProfileId)))
			objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,rsGetUserID.Fields.Item("customer_ID").Value))
			objCmd.Execute()
		
	End If 'if auth.net response is a success for creating customer CIM ID
	
	' if customer opted to save shipping/billing information add it to auth.net CIM below
	if request.form("save-all") = "on" then
	
		' Add a new SHIPPING address
		strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<createCustomerShippingAddressRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <customerProfileId>" & strCustomerProfileId & "</customerProfileId>" _
		& "<address>" _
		& "  <firstName>" & request.form("shipping-first") & "</firstName>" _
		& "  <lastName>" & request.form("shipping-last") & "</lastName>" _
		& "  <company>" & request.form("shipping-company") & "</company>" _
		& "  <address>" & request.form("shipping-address") & "|" & request.form("shipping-address2") & "</address>" _
		& "  <city>" & request.form("shipping-city") & "</city>" _
		& "  <state>" & request.form("shipping-state") & "|" & request.form("shipping-province-canada") & "" & request.form("shipping-province") & "</state>" _
		& "  <zip>" & request.form("shipping-zip") & "</zip>" _
		& "  <country>" & request.form("shipping-country") & "</country>" _
		& "  <phoneNumber>" & request.form("shipping-phone") & "</phoneNumber>" _
		& "</address>" _
		& "</createCustomerShippingAddressRequest>"
		
		
		Set objResponseAddShipping = SendApiRequest(strReq)

		' If succcess in adding shipping address to CIM also store authorize.net ID for address in our address book table
		If IsApiResponseSuccess(objResponseAddShipping) Then

			strCustomerShippingId = objResponseAddShipping.selectSingleNode("/*/api:customerAddressId").Text
			
			'for use on checkout/inc_save_cims_to_order.asp
			var_cim_shipping_id = objResponseAddShipping.selectSingleNode("/*/api:customerAddressId").Text
				
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO TBL_AddressBook (cim_shippingid, custID, address_type, default_shipping) VALUES (?,?,?,?)"
			objCmd.Parameters.Append(objCmd.CreateParameter("cim_shippingid",200,1,30,Server.HTMLEncode(strCustomerShippingId)))
			objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,rsGetUserID.Fields.Item("customer_ID").Value))
			objCmd.Parameters.Append(objCmd.CreateParameter("address_type",200,1,30,"shipping"))
			objCmd.Parameters.Append(objCmd.CreateParameter("shipping_default",3,1,10,1)) ' 1 default, 0 not default
			objCmd.Execute()
		
		End if
		
		' ADD NEW PAYMENT / BILLING address
		strReqBillingAdd = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<createCustomerPaymentProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <customerProfileId>" & strCustomerProfileId & "</customerProfileId>" _
		& "<paymentProfile>" _
		& "<billTo>" _
		& "  <firstName>" & Request.Form("billing-first") & "</firstName>" _
		& "  <lastName>" & Request.Form("billing-last") & "</lastName>" _
		& "  <address>" & Request.Form("billing-address") & "|" & Request.Form("billing-address2") & "</address>" _
		& "  <city>" & Request.Form("billing-city") & "</city>" _
		& "  <state>" & Request.Form("billing-state") & "|" & Request.Form("billing-province-canada") & "" & request.form("billing-province") & "</state>" _
		& "  <zip>" & Request.Form("billing-zip") & "</zip>" _
		& "  <country>" & Request.Form("billing-country") & "</country>" _
		& "</billTo>" _
		& "<payment>" _
		& "<creditCard>" _
		& "  <cardNumber>" & Replace(Request.Form("card_number"), " ", "") & "</cardNumber>" _
		& "  <expirationDate>" & Request.Form("billing-month") & "" & Request.Form("billing-year") & "</expirationDate>" _
		& "</creditCard>" _
		& "</payment>" _
		& "</paymentProfile>" _
		& "</createCustomerPaymentProfileRequest>"
		
		
		Set objResponseAddBilling = SendApiRequest(strReqBillingAdd)

			' If succcess in adding billing address to CIM also store authorize.net ID for address in our address book table
			If IsApiResponseSuccess(objResponseAddBilling) Then
			
				update_successful = "yes"
				
				strCustomerBillingId = objResponseAddBilling.selectSingleNode("/*/api:customerPaymentProfileId").Text
				
				'for use on checkout/inc_save_cims_to_order.asp
				var_cim_billing_id = objResponseAddBilling.selectSingleNode("/*/api:customerPaymentProfileId").Text
			
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "INSERT INTO TBL_AddressBook (cim_shippingid, custID, address_type, default_billing) VALUES (?,?,?,?)"
				objCmd.Parameters.Append(objCmd.CreateParameter("cim_shippingid",200,1,30,Server.HTMLEncode(strCustomerBillingId)))
				objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,rsGetUserID.Fields.Item("customer_ID").Value))
				objCmd.Parameters.Append(objCmd.CreateParameter("address_type",200,1,30,"billing"))
				objCmd.Parameters.Append(objCmd.CreateParameter("billing_default",3,1,10,1))
				objCmd.Execute()
			
			End if
		' ADD NEW PAYMENT / BILLING address
		
	
	end if ' if customer opted to save shipping/billing information add it to auth.net CIM below


	var_create_account_status = "done"

%>	
	<!--#include virtual="/emails/function-send-email.asp"-->
	<!--#include virtual="/emails/email_variables.asp"-->
	<!--#include virtual="/accounts/inc_transfer_cart_contents.asp" -->
<%
	
	' clear all DB and variables
	set rsGetUserID = nothing
	
end if ' if password is not empty
' END create new customer account
' =================================================================================
%>