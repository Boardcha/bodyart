<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/Connections/authnet.asp"-->
<%
' Pull the customer information from a cookie
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM customers  WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
		Set rsGetUser = objCmd.Execute()

' Delete SHIPPING address
if request.form("type") = "shipping" then

		strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<deleteCustomerShippingAddressRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <customerProfileId>" & rsGetUser.Fields.Item("cim_custid").Value & "</customerProfileId>" _
		& "  <customerAddressId>" & request.form("id") & "</customerAddressId>" _
		& "</deleteCustomerShippingAddressRequest>"
		
		
		Set objResponseAddShipping = SendApiRequest(strReq)

		' If succcess deleteing address on Auth.net CIM then delete the row from our local database as well
		If IsApiResponseSuccess(objResponseAddShipping) Then
		

		
		End if

		' Force it to delete from our database
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "DELETE FROM TBL_AddressBook WHERE custID = ? AND cim_shippingid = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("cim_custid",200,1,30,CustID_Cookie))
		objCmd.Parameters.Append(objCmd.CreateParameter("shipping_id",3,1,10,request.form("id")))
		objCmd.Execute()
			
%>
{
	"status":"success",
	"status_text":"Address has been removed"
}
<%

end if 
' End delete SHIPPING address


' Delete BILLING CREDIT CARD
if request.form("type") = "billing" then

		strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<deleteCustomerPaymentProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <customerProfileId>" & rsGetUser.Fields.Item("cim_custid").Value & "</customerProfileId>" _
		& "  <customerPaymentProfileId>" & request.form("id") & "</customerPaymentProfileId>" _
		& "</deleteCustomerPaymentProfileRequest>"
		Set objResponseDeleteBilling = SendApiRequest(strReq)

		' If succcess deleteing address on Auth.net CIM then delete the row from our local database as well
		If IsApiResponseSuccess(objResponseDeleteBilling) Then
		
		
		End if

		' Force it to delete from our database
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "DELETE FROM TBL_AddressBook WHERE custID = ? AND cim_shippingid = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("cim_custid",200,1,30,CustID_Cookie))
		objCmd.Parameters.Append(objCmd.CreateParameter("shipping_id",3,1,10,request.form("id")))
		objCmd.Execute()
			
%>
{
	"status":"success",
	"status_text":"Credit card has been removed"
}
<%

end if ' End delete BILLING CREDIT CARD

DataConn.Close()
Set DataConn = Nothing
%>
