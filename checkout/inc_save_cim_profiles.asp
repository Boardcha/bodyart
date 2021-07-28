<%
' ADD A NEW BILLING ADDRESS PROFILE ======================================================================
If CustID_Cookie <> "" and CustID_Cookie <> 0 Then ' if customer is logged in

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT cim_custid FROM customers WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
	Set rsGetUser = objCmd.Execute()
	
	' Set customers main CIM ID to variable
	var_cim_custid = rsGetUser.Fields.Item("cim_custid").Value

If Request.Form("billing-first") <> "" AND Request.Form("billing-last") <> "" AND request.form("card-save") <> "" AND request.form("billing-status") = "add" Then

		strReqBillingAdd = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<createCustomerPaymentProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <customerProfileId>" & var_cim_custid & "</customerProfileId>" _
		& "<paymentProfile>" _
		& "<billTo>" _
		& "  <firstName>" & Request.Form("billing-first") & "</firstName>" _
		& "  <lastName>" & Request.Form("billing-last") & "</lastName>" _
		& "  <company>" & Request.Form("billing-company") & "</company>" _
		& "  <address>" & Request.Form("billing-address") & "|" & Request.Form("billing-address2") & "</address>" _
		& "  <city>" & Request.Form("billing-city") & "</city>" _
		& "  <state>" & Request.Form("billing-state") & "|" & Request.Form("billing-province") & "" & Request.Form("billing-province-canada") & "</state>" _
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
		
			strCustomerBillingId = objResponseAddBilling.selectSingleNode("/*/api:customerPaymentProfileId").Text

			' used to store into order on checkout/inc_save_cims_to_order
			var_cim_billing_id = strCustomerBillingId
			
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO TBL_AddressBook (cim_shippingid, custID, address_type, default_billing) VALUES (?,?,?,?)"
			objCmd.Parameters.Append(objCmd.CreateParameter("cim_shippingid",200,1,30,Server.HTMLEncode(strCustomerBillingId)))
			objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
			objCmd.Parameters.Append(objCmd.CreateParameter("address_type",200,1,30,"billing"))
			objCmd.Parameters.Append(objCmd.CreateParameter("billing_default",3,1,10,0))
			objCmd.Execute()
		
		End if

End If
' ADD A NEW BILLING ADDRESS PROFILE ======================================================================


' ADD A NEW SHIPPING ADDRESS PROFILE ======================================================================
If Request.Form("shipping-first") <> "" AND Request.Form("shipping-last") <> "" AND request.form("shipping-save") <> "" AND request.form("shipping-status") = "add"  Then

		strReqShippingAdd = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<createCustomerShippingAddressRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <customerProfileId>" & var_cim_custid & "</customerProfileId>" _
		& "<address>" _
		& "  <firstName>" & Request.Form("shipping-first") & "</firstName>" _
		& "  <lastName>" & Request.Form("shipping-last") & "</lastName>" _
		& "  <company>" & Request.Form("shipping-company") & "</company>" _
		& "  <address>" & Request.Form("shipping-address") & "|" & Request.Form("shipping-address2") & "</address>" _
		& "  <city>" & Request.Form("shipping-city") & "</city>" _
		& "  <state>" & Request.Form("shipping-state") & "|" & Request.Form("shipping-province") & "" & Request.Form("shipping-province-canada") & "</state>" _
		& "  <zip>" & Request.Form("shipping-zip") & "</zip>" _
		& "  <country>" & Request.Form("shipping-country") & "</country>" _
		& "  <phoneNumber>" & Request.Form("shipping-phone") & "</phoneNumber>" _
		& "</address>" _
		& "</createCustomerShippingAddressRequest>"
		Set objResponseAddShipping = SendApiRequest(strReqShippingAdd)

		' If succcess in adding shipping address to CIM also store authorize.net ID for address in our address book table
		If IsApiResponseSuccess(objResponseAddShipping) Then
			
			strCustomerShippingId = objResponseAddShipping.selectSingleNode("/*/api:customerAddressId").Text

			' used to store into order on checkout/inc_save_cims_to_order
			var_cim_shipping_id = strCustomerShippingId
	
			
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO TBL_AddressBook (cim_shippingid, custID, address_type, default_shipping) VALUES (?,?,?,?)"
			objCmd.Parameters.Append(objCmd.CreateParameter("cim_shippingid",200,1,30,Server.HTMLEncode(strCustomerShippingId)))
			objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
			objCmd.Parameters.Append(objCmd.CreateParameter("address_type",200,1,30,"shipping"))
			objCmd.Parameters.Append(objCmd.CreateParameter("shipping_default",3,1,10,0))
			objCmd.Execute()
		
		End if
		

End If
' ADD A NEW SHIPPING ADDRESS PROFILE ======================================================================

end if  ' if customer is logged in
%>