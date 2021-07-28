<%
' UPDATE A BILLING ADDRESS PROFILE ======================================================================
If request.form("billing-status") = "update" Then

	if Request.Form("billing-country") = "USA" then
		var_billing_state = Request.Form("billing-state") & "|"
	elseif  Request.Form("billing-country") = "Canada" then
		var_billing_state = "|" & Request.Form("billing-province-canada")
	else
		var_billing_state = "|" & Request.Form("billing-province")
	end if

		strReqUpdateBilling = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<updateCustomerPaymentProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <customerProfileId>" & var_cim_custid & "</customerProfileId>" _
		& "<paymentProfile>" _
		& "<billTo>" _
		& "  <firstName>" & Request.Form("billing-first") & "</firstName>" _
		& "  <lastName>" & Request.Form("billing-last") & "</lastName>" _
		& "  <company>" & Request.Form("billing-company") & "</company>" _
		& "  <address>" & Request.Form("billing-address") & "|" & Request.Form("billing-address2") & "</address>" _
		& "  <city>" & Request.Form("billing-city") & "</city>" _
		& "  <state>" & var_billing_state & "</state>" _
		& "  <zip>" & Request.Form("billing-zip") & "</zip>" _
		& "  <country>" & Request.Form("billing-country") & "</country>" _
		& "</billTo>" _
		& "<payment>" _
		& "<creditCard>" _
		& "  <cardNumber>" & Replace(Request.Form("card_number"), " ", "") & "</cardNumber>" _
		& "  <expirationDate>" & Request.Form("billing-month") & "" & Request.Form("billing-year") & "</expirationDate>" _
		& "</creditCard>" _
		& "</payment>" _
		& "<customerPaymentProfileId>" & request.form("cim_billing") & "</customerPaymentProfileId>" _
		& "</paymentProfile>" _
		& "</updateCustomerPaymentProfileRequest>"		
		Set objResponseUpdateBilling = SendApiRequest(strReqUpdateBilling)
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE TBL_AddressBook SET last_updated = '" & now() & "' WHERE cim_shippingid = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,12,	request.form("cim_billing")))
		objCmd.Execute()

End If
' UPDATE A BILLING ADDRESS PROFILE ======================================================================


' UPDATE A SHIPPING ADDRESS PROFILE ======================================================================
If request.form("shipping-status") = "update" Then

	if Request.Form("shipping-country") = "USA" then
		var_shipping_state = Request.Form("shipping-state") & "|"
	elseif  Request.Form("shipping-country") = "Canada" then
		var_shipping_state = "|" & Request.Form("shipping-province-canada")
	else
		var_shipping_state = "|" & Request.Form("shipping-province")
	end if
	
		strReqShippingUpdate = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<updateCustomerShippingAddressRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <customerProfileId>" & var_cim_custid & "</customerProfileId>" _
		& "<address>" _
		& "  <firstName>" & Request.Form("shipping-first") & "</firstName>" _
		& "  <lastName>" & Request.Form("shipping-last") & "</lastName>" _
		& "  <company>" & Request.Form("shipping-company") & "</company>" _
		& "  <address>" & Request.Form("shipping-address") & "|" & Request.Form("shipping-address2") & "</address>" _
		& "  <city>" & Request.Form("shipping-city") & "</city>" _
		& "  <state>" & var_shipping_state & "</state>" _
		& "  <zip>" & Request.Form("shipping-zip") & "</zip>" _
		& "  <country>" & Request.Form("shipping-country") & "</country>" _
		& "  <phoneNumber>" & Request.Form("shipping-phone") & "</phoneNumber>" _
		& "  <customerAddressId>" & request.form("cim_shipping") & "</customerAddressId>" _
		& "</address>" _
		& "</updateCustomerShippingAddressRequest>"		
		Set objResponseUpdateShipping = SendApiRequest(strReqShippingUpdate)
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE TBL_AddressBook SET last_updated = '" & now() & "' WHERE cim_shippingid = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,12,	request.form("cim_shipping")))
		objCmd.Execute()

End If
' UPDATE A SHIPPING ADDRESS PROFILE ======================================================================
%>