<%
If request.form("googlepay") = "on" Then
	' Connect to auth.net
	strChargeCard = "<?xml version=""1.0"" encoding=""utf-8""?>" _
	& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
	& MerchantAuthentication() _
	& "<transactionRequest>" _
		& "<transactionType>authCaptureTransaction</transactionType>" _
		& "<amount>" & request.form("amount") & "</amount>" _	
		& "<payment>" _
			& "<opaqueData>" _
				& "<dataDescriptor>COMMON.GOOGLE.INAPP.PAYMENT</dataDescriptor>" _
				& "<dataValue>" & request.form("encryptedToken") & "</dataValue>" _
			& "</opaqueData>" _
		& "</payment>" _  
		& "<order>" _
		&   "<invoiceNumber>" & Session("invoiceid") & "</invoiceNumber>" _
		&   "<description>Body jewelry</description>" _
		& "</order>" _		
		& "<tax>" _  
		&   "<amount>" & request.form("tax") & "</amount>" _  
		&   "<name>Tax</name>" _  
		&   "<description></description>" _  
		& "</tax>" _  
		& "<shipping>" _  
		&   "<amount>" & request.form("shipping_amount") & "</amount>" _  
		&   "<name>Shipping</name>" _  
		&   "<description></description>" _  
		& "</shipping>" _  
		& "<billTo>" _  
		&   "<firstName>" & request.form("full_name") & "</firstName>" _  
		&   "<lastName></lastName>" _  
		&   "<address>" & request.form("address1") & " " & request.form("address2") & "</address>" _  
		&   "<city>" & request.form("locality") & "</city>" _  
		&   "<state>" & request.form("administrative_area") & "</state>" _  
		&   "<zip>" & request.form("postal_code") & "</zip>" _  
		&   "<country>" & request.form("country_code") & "</country>" _  
		& "</billTo>" _  
	& "</transactionRequest>" _
	& "</createTransactionRequest>"

	Set objResponseChargeCard = SendApiRequest(strChargeCard)

	If IsApiResponseSuccess(objResponseChargeCard) Then
		strTransactionId = objResponseChargeCard.selectSingleNode("/*/api:transactionResponse/api:transId").Text
		session("cc_transid") = strTransactionId
		strCardType = objResponseChargeCard.selectSingleNode("/*/api:transactionResponse/api:accountType").Text
		
		' If approved... ' 1 = Approved, 2 = Declined, 3 = Error, 4 = Held for Review
		If objResponseChargeCard.selectSingleNode("/*/api:transactionResponse/api:responseCode").Text = 1 Then 
			%>
				"cc_approved":"yes", "cc_reason":"approved",
				"google_pay_response":"<%=objResponseChargeCard.selectSingleNode("/*/api:transactionResponse").Text%>"
			<% 	
			payment_approved = "yes"
			mailer_type = "cc approved"
			session("cc_status") = "approved"
		Else
			%>
				"cc_approved":"no", "cc_reason":"<%= objResponseChargeCard.selectSingleNode("/*/api:transactionResponse/api:errors/api:error/api:errorText").Text %>",
				"google_pay_response":"<%=objResponseChargeCard.selectSingleNode("/*/api:transactionResponse").Text%>"
			<% 			
			payment_approved = "no"	
			session("cc_status") = "declined"
			session("cc_decline_reason") = objResponseChargeCard.selectSingleNode("/*/api:transactionResponse/api:errors/api:error/api:errorText").Text			
		End If
	Else
		%>
			"cc_approved":"no", "cc_reason":"Problem with payment information",
			"google_pay_response":"<%=objResponseChargeCard.selectSingleNode("/*/api:transactionResponse").Text%>"
		<% 			
	End If


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
End If	
%>