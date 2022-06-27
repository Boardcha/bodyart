<%@LANGUAGE="VBSCRIPT" CodePage = 65001 %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/Connections/authnet.asp"-->
<%
firstname = Request.Form("first")
lastname = Request.Form("last")
If Request.Form("shipping-address") <> "" Then address =  Request.Form("shipping-address") & "|" & Request.Form("shipping-address2") Else address =  Request.Form("address") & "|" & Request.Form("address2")
If Request.Form("shipping-city") <> "" Then city = Request.Form("shipping-city") Else city = Request.Form("city")
If Request.Form("shipping-state") <> "" Or Request.Form("shipping-province") <> "" Or Request.Form("shipping-province-canada") <> "" Then
	state = Request.Form("shipping-state") & "|" & Request.Form("shipping-province") & Request.Form("shipping-province-canada") 
Else
	state = Request.Form("state") & "|" & Request.Form("province") & Request.Form("province-canada") 
End If
If Request.Form("shipping-zip") <> "" Then zip = Request.Form("shipping-zip") Else zip = Request.Form("zip")
If Request.Form("shipping-country") <> "" Then country = Request.Form("shipping-country") Else country = Request.Form("country")
If Request.Form("shipping-phone") <> "" Then phone = Request.Form("shipping-phone") Else phone = Request.Form("phone")

' Pull the customer information from a cookie
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM customers  WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
	Set rsGetUser = objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_AddressBook SET last_updated = '" & now() & "' WHERE cim_shippingid = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,12,request.form("cim-id")))
	objCmd.Execute()
	


' Update a new SHIPPING address
If request.form("type") = "shipping" Then

		strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<updateCustomerShippingAddressRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <customerProfileId>" & rsGetUser.Fields.Item("cim_custid").Value & "</customerProfileId>" _
		& "<address>" _
		& "  <firstName>" & firstname & "</firstName>" _
		& "  <lastName>" & lastname & "</lastName>" _
		& "  <company></company>" _
		& "  <address>" & address & "</address>" _
		& "  <city>" & city & "</city>" _
		& "  <state>" & state & "</state>" _
		& "  <zip>" & zip & "</zip>" _
		& "  <country>" & country & "</country>" _
		& "  <phoneNumber>" & phone & "</phoneNumber>" _
		& "  <customerAddressId>" & request.form("cim-id") & "</customerAddressId>" _
		& "</address>" _
		& "</updateCustomerShippingAddressRequest>"		
		Set objResponseUpdateShipping = SendApiRequest(strReq)
		
		If IsApiResponseSuccess(objResponseUpdateShipping) Then
%>
{
	"status":"success"
}
<%		
		else
%>
{
	"status":"fail"
}
<%		
		
		End if

End If ' Update add new SHIPPING address




' Update BILLING CREDIT CARD information
If request.form("type") = "billing" Then



		strReqUpdateBilling = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<updateCustomerPaymentProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <customerProfileId>" & rsGetUser.Fields.Item("cim_custid").Value & "</customerProfileId>" _
		& "<paymentProfile>" _
		& "<billTo>" _
		& "  <firstName>" & firstname & "</firstName>" _
		& "  <lastName>" & lastname & "</lastName>" _
		& "  <company></company>" _
		& "  <address>" & address & "</address>" _
		& "  <city>" & city & "</city>" _
		& "  <state>" & state & "</state>" _
		& "  <zip>" & zip & "</zip>" _
		& "  <country>" & country & "</country>" _
		& "</billTo>" _
		& "<payment>" _
		& "<creditCard>" _
		& "  <cardNumber>" & Replace(Request.Form("card_number"), " ", "") & "</cardNumber>" _
		& "  <expirationDate>" & Request.Form("billing-month") & "" & Request.Form("billing-year") & "</expirationDate>" _
		& "</creditCard>" _
		& "</payment>" _
		& "<customerPaymentProfileId>" & request.form("cim-id") & "</customerPaymentProfileId>" _
		& "</paymentProfile>" _
		& "</updateCustomerPaymentProfileRequest>"		
		Set objResponseUpdateBilling = SendApiRequest(strReqUpdateBilling)
		
		If IsApiResponseSuccess(objResponseUpdateBilling) Then
%>
{
	"status":"success"
}
<%		
		else
%>
{
	"status":"fail",
	"reason": "<% PrintErrors(objResponseUpdateBilling) %>"
}
<%
		End if		
		

End If ' BILLING CREDIT CARD information




DataConn.Close()
Set DataConn = Nothing
%>
