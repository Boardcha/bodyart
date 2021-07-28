<%
' BILL AUTH.NET CIM IF CUSTOMER IS REGISTERED ==========================================================
if var_grandtotal > 0 and request.form("paypal") <> "on" and request.form("afterpay") <> "on" and request.form("cash") <> "on" and request.form("cim_billing") <> "paypal" and request.form("cim_billing") <> "cash" then

if request.form("cim_billing") <> "" then

	var_billTo_cim = "<profile><customerProfileId>" & session("cim_accountNumber") & "</customerProfileId><paymentProfile><paymentProfileId>" & request.form("cim_billing") & "</paymentProfileId></paymentProfile></profile>"

else

	var_billTo_fields = "<billTo><firstName>" & request.form("billing-first") & "</firstName><lastName>" & request.form("billing-last") & "</lastName><address>" & Request.Form("billing-address") & "|" & Request.Form("billing-address2") & "</address><city>" & request.form("billing-city") & "</city><state>" & Request.Form("billing-state") & "|" & Request.Form("billing-province") & "" & Request.Form("billing-province-canada") & "</state><zip>" & Request.Form("billing-zip") & "</zip><country>" & request.form("billing-country") & "</country></billTo>"
	
		
	if request.form("cvv2") <> "" then
		var_cvv2 = "<cardCode>" & request.form("cvv2") & "</cardCode>"
	end if

	var_cc_fields = "<payment><creditCard><cardNumber>" & replace(request.form("card_number"), " ", "") & "</cardNumber><expirationDate>" & request.form("billing-month") & request.form("billing-year") & "</expirationDate>" & var_cvv2 & "</creditCard></payment>"
	
end if

		' Connect to auth.net
		strChargeCard = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "<transactionRequest>" _
		& "		<transactionType>authCaptureTransaction</transactionType>" _
		& "<amount>" & FormatNumber(var_grandtotal,,,,0) & "</amount>" _
		& var_cc_fields _ 
		& var_billTo_cim _
		& "<order>" _
		& "<invoiceNumber>" & Session("invoiceid") & "</invoiceNumber>" _
		& "<description>Body jewelry</description>" _
		& "</order>" _
		& var_billTo_fields _
		& "</transactionRequest>" _
		& "</createTransactionRequest>"
		Set objResponseChargeCard = SendApiRequest(strChargeCard)

		
		
		' APPROVED - If REGISTERED customer order is APPROVED -----------------------------------
		If IsApiResponseSuccess(objResponseChargeCard) Then
		
		' set variables for approved OR declined responses
		strTransactionId = objResponseChargeCard.selectSingleNode("/*/api:transactionResponse/api:transId").Text
		
		session("cc_transid") = strTransactionId
		
		strCardType = objResponseChargeCard.selectSingleNode("/*/api:transactionResponse/api:accountType").Text
		
		' 1 = Approved, 2 = Declined, 3 = Error, 4 = Held for Review
		if objResponseChargeCard.selectSingleNode("/*/api:transactionResponse/api:responseCode").Text = 1 then ' if approved
		
%>
	"cc_approved":"yes",
	"cc_reason":"approved"
<% 
			' Set variable for use on other includes
			payment_approved = "yes"
			mailer_type = "cc approved"
			session("cc_status") = "approved"
		'	strCardType = "Credit card"
			
		
		else ' payment declined
						
				
%>
	"cc_approved":"no",
	"cc_reason":"<%= objResponseChargeCard.selectSingleNode("/*/api:transactionResponse/api:errors/api:error/api:errorText").Text %>"
<% 			
			payment_approved = "no"	
			session("cc_status") = "declined"
			session("cc_decline_reason") = objResponseChargeCard.selectSingleNode("/*/api:transactionResponse/api:errors/api:error/api:errorText").Text
			
		end if ' if payment is declined	
		
		else ' Also set a decline if no results are returned from auth.net

%>
	"cc_approved":"no",
	"cc_reason":"Problem with payment information"
<% 		
			payment_approved = "no"	
			session("cc_status") = "declined"
			session("cc_decline_reason") = "Problem with payment information"
			strCardType = "Credit card"
			
		end if ' if response came back from Auth.net
		

	' ---- Add transaction ID, and response verification information to order
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.NamedParameters = True
	objCmd.CommandText = "UPDATE sent_items SET transactionID = ?, pay_method = ? WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("@TransactionID", 200,1,30, strTransactionId))
	objCmd.Parameters.Append(objCmd.CreateParameter("@strCardType", 200,1,30, strCardType))
	objCmd.Parameters.Append(objCmd.CreateParameter("@InvoiceID", 3,1,10, Session("invoiceid")))
	objCmd.Execute()
	set objCmd = Nothing

end if ' make sure it's not a paypal or cash order
%>