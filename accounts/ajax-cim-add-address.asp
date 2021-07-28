<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/Connections/authnet.asp"-->

<%
' Pull the customer information from a cookie
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM customers  WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
	Set rsGetUser = objCmd.Execute()
	
For Each item In Request.Form
'	Response.Write "Key: " & item & " - Value: " & Request.Form(item) & "<BR />"
Next	

' Add a new SHIPPING address
If request.form("type") = "shipping" Then

	strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
	& "<createCustomerShippingAddressRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
	& MerchantAuthentication() _
	& "  <customerProfileId>" & rsGetUser.Fields.Item("cim_custid").Value & "</customerProfileId>" _
	& "<address>" _
	& "  <firstName>" & Request.Form("first") & "</firstName>" _
	& "  <lastName>" & Request.Form("last") & "</lastName>" _
	& "  <company>" & Request.Form("company") & "</company>" _
	& "  <address>" & Request.Form("address") & "|" & Request.Form("address2") & "</address>" _
	& "  <city>" & Request.Form("city") & "</city>" _
	& "  <state>" & Request.Form("state") & "|" & Request.Form("province") & Request.Form("province-canada") & "</state>" _
	& "  <zip>" & Request.Form("zip") & "</zip>" _
	& "  <country>" & Request.Form("country") & "</country>" _
	& "  <phoneNumber>" & Request.Form("phone") & "</phoneNumber>" _
	& "</address>" _
	& "</createCustomerShippingAddressRequest>"
	
	
	Set objResponseAddShipping = SendApiRequest(strReq)

	' If succcess in adding shipping address to CIM also store authorize.net ID for address in our address book table
	If IsApiResponseSuccess(objResponseAddShipping) Then
	
		strCustomerShippingId = objResponseAddShipping.selectSingleNode("/*/api:customerAddressId").Text
	
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO TBL_AddressBook (cim_shippingid, custID, address_type) VALUES (?,?,?)"
		objCmd.Parameters.Append(objCmd.CreateParameter("cim_shippingid",200,1,30,Server.HTMLEncode(strCustomerShippingId)))
		objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,rsGetUser.Fields.Item("customer_ID").Value))
		objCmd.Parameters.Append(objCmd.CreateParameter("address_type",200,1,30,"shipping"))
		objCmd.Execute()
%>
{
	"status":"success"
}
<%		
		else

		var_message = objResponseAddShipping.selectSingleNode("/*/api:messages/api:message/api:text").Text
%>
{
	"status":"fail",
	"message": "<%= var_message %>"
}
<%		
		
		End if		

		

End If ' End add new SHIPPING address




' Add a new BILLING address
If request.form("type") = "billing" Then

		strReqBillingAdd = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<createCustomerPaymentProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <customerProfileId>" & rsGetUser.Fields.Item("cim_custid").Value & "</customerProfileId>" _
		& "<paymentProfile>" _
		& "<billTo>" _
		& "  <firstName>" & Request.Form("first") & "</firstName>" _
		& "  <lastName>" & Request.Form("last") & "</lastName>" _
		& "  <company>" & Request.Form("company") & "</company>" _
		& "  <address>" & Request.Form("address") & "|" & Request.Form("address2") & "</address>" _
		& "  <city>" & Request.Form("city") & "</city>" _
		& "  <state>" & Request.Form("state") & "|" & Request.Form("province") & Request.Form("province-canada") & "</state>" _
		& "  <zip>" & Request.Form("zip") & "</zip>" _
		& "  <country>" & Request.Form("country") & "</country>" _
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
				
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO TBL_AddressBook (cim_shippingid, custID, address_type) VALUES (?,?,?)"
			objCmd.Parameters.Append(objCmd.CreateParameter("cim_shippingid",200,1,30,Server.HTMLEncode(strCustomerBillingId)))
			objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,rsGetUser.Fields.Item("customer_ID").Value))
			objCmd.Parameters.Append(objCmd.CreateParameter("address_type",200,1,30,"billing"))
			objCmd.Execute()
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

End If ' End add new BILLING address




DataConn.Close()
Set DataConn = Nothing
%>
