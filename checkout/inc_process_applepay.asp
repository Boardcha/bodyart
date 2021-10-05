<%
If request.form("applepay") = "on" Then
	' Connect to auth.net
	strChargeCard = "<?xml version=""1.0"" encoding=""utf-8""?>" _
	& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
	& MerchantAuthentication() _
	& "<transactionRequest>" _
		& "<transactionType>authCaptureTransaction</transactionType>" _
		& "<amount>" & request.form("amount") & "</amount>" _	
		& "<payment>" _
			& "<opaqueData>" _
				& "<dataDescriptor>COMMON.APPLE.INAPP.PAYMENT</dataDescriptor>" _
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
		&   "<address>" & request.form("address") & "</address>" _  
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
				"apple_pay_response":"<%=objResponseChargeCard.selectSingleNode("/*/api:transactionResponse").Text%>"
			<% 	
		Else
			%>
				"cc_approved":"no", "cc_reason":"<%= objResponseChargeCard.selectSingleNode("/*/api:transactionResponse/api:errors/api:error/api:errorText").Text %>",
				"apple_pay_response":"<%=objResponseChargeCard.selectSingleNode("/*/api:transactionResponse").Text%>"
			<% 			
		End If
	Else
		%>
			"cc_approved":"no", "cc_reason":"Problem with payment information",
			"apple_pay_response":"<%=objResponseChargeCard.selectSingleNode("/*/api:transactionResponse").Text%>"
		<% 			
	End If

	Response.Write "AUTH.NET RESPONSE:<br>" & objResponseChargeCard.selectSingleNode("/*/api:transactionResponse").Text & "<br><br>"	
	
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