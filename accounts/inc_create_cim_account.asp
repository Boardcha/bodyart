<%
' CREATE NEW CIM ACCOUNT ===========================================================================
if CustID_Cookie <> "" then ' if customer is logged in create a CIM account

	'cust id variable set on "inc_save_order.asp" page
	strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
	& "<createCustomerProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
	& MerchantAuthentication() _
	& "<profile>" _
	& "  <merchantCustomerId>" & CustID_Cookie & "</merchantCustomerId>" _
	& "  <email>" & var_email & "</email>" _
	& "</profile>" _
	& "</createCustomerProfileRequest>"
	Set objResponseCreateProfile = SendApiRequest(strReq)

	' If succcess in created a new CIM profile ID then add that new ID to our database in the customers table
	If IsApiResponseSuccess(objResponseCreateProfile) Then
	  strCustomerProfileId = objResponseCreateProfile.selectSingleNode("/*/api:customerProfileId").Text
		
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "UPDATE customers SET cim_custid = ? WHERE customer_ID = ?"
			objCmd.Parameters.Append(objCmd.CreateParameter("cim_custid",200,1,30,Server.HTMLEncode(strCustomerProfileId)))
			objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10, CustID_Cookie))
			objCmd.Execute()
	else
	'	Response.Write "The operation failed with the following errors:<br>" & vbCrLf
	'	PrintErrors(objResponseCreateProfile)
	End If
	
end if 
' CREATE NEW CIM ACCOUNT ===========================================================================

%>