<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/Connections/authnet.asp"-->
<%
	' decrypt refund information
	Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
	password = "3uBRUbrat77V"
	data = request.querystring("encrypted")
	decrypted_refund = objCrypt.Decrypt(password, data)
	
	split_refund = split(decrypted_refund, "|")

	invoice_id = split_refund(0)
    ProductDetailID = split_refund(1)
	var_customer_number = split_refund(2)

    var_refund_id = request.querystring("id")

	Set objCrypt = Nothing

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT REF.*, sent_items.customer_id, email, customer_first, transactionID, pay_method FROM sent_items INNER JOIN TBL_Refunds_backordered_items REF ON sent_items.ID = REF.invoice_id WHERE redeemed = 0 AND REF.invoice_id = ? AND REF.ProductDetailID = ? AND REF.encrypted_code = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, invoice_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,15, ProductDetailID))
	objCmd.Parameters.Append(objCmd.CreateParameter("encrypted_code",200,1,200, data))
	set rsCheckRefund = objCmd.Execute()


    if not rsCheckRefund.eof then
	
        var_db_refund_amt = formatnumber(rsCheckRefund.Fields.Item("refund_total").Value,2)

        '========== GET CARD INFORMATION FROM TRANSACTION ==========================
        strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
        & "<getTransactionDetailsRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
        & MerchantAuthentication() _
        & "<transId>" & rsCheckRefund.Fields.Item("transactionID").Value & "</transId>" _
        & "</getTransactionDetailsRequest>"
    
        Set objGetTransactionDetails = SendApiRequest(strReq)
    
        ' If succcess retrieve transaction information
        If IsApiResponseSuccess(objGetTransactionDetails) Then
            If not(objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:payment/api:creditCard/api:cardNumber") is nothing) then
                strCardNumber = objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:payment/api:creditCard/api:cardNumber").Text
            end if     
        Else ' if there's an error getting a transaction
            Response.Write "The operation failed with the following errors:<br>" & vbCrLf
            PrintErrors(objGetTransactionDetails)
        End If

        ' ====== PROCESS THE REFUND THROUGH AUTHORIZE.NET  =========================  
        if  rsCheckRefund.Fields.Item("pay_method").Value <> "PayPal" then
            var_card_info = "<payment><creditCard><cardNumber>" & strCardNumber & "</cardNumber><expirationDate>XXXX</expirationDate></creditCard></payment>"
        else
            var_card_info = ""
        end if
	 
		strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "<transactionRequest>" _
		& "		<transactionType>refundTransaction</transactionType>" _
		& "		<amount>" & var_db_refund_amt & "</amount>" _
				& var_card_info _
		& "		<refTransId>" & rsCheckRefund.Fields.Item("transactionID").Value & "</refTransId>" _
		& "		<order>" _
		& "			<invoiceNumber>" & invoice_id & "</invoiceNumber>" _
		& "			<description>Order refund</description>" _
		& "		</order>" _
		& "</transactionRequest>" _
		& "</createTransactionRequest>"
		

		Set objResponse = SendApiRequest(strSend)
	
		var_message = objResponse.selectSingleNode("/*/api:messages/api:message/api:text").Text

		'=======  APPROVED ORDER ====================================================
		If IsApiResponseSuccess(objResponse) Then
	
			var_responseCode = objResponse.selectSingleNode("/*/api:transactionResponse/api:responseCode").Text
	
			if var_responseCode = 1 then ' approved 
	
				' ==== Set item's backorder status ====
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE TBL_OrderSummary SET backorder = 0, BackorderReview = 'N', notes = 'Customer has refunded the item and backorder has been cleared' WHERE InvoiceID = ? AND DetailID = ?"
				objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,15, invoice_id))
				objCmd.Parameters.Append(objCmd.CreateParameter("detailID",3,1,20, ProductDetailID))
				objCmd.Execute()
				
				' ====== update the record to clear it out =======
				set objCmd = Server.CreateObject("ADODB.Command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE TBL_Refunds_backordered_items redeemed = 1, date_redeemed = GETDATE(), redeemedAs = 'Refund' WHERE invoice_id = ? AND encrypted_code = ? AND id = ?"
				objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, invoice_id))
				objCmd.Parameters.Append(objCmd.CreateParameter("encrypted_code",200,1,200, data))
				objCmd.Parameters.Append(objCmd.CreateParameter("var_refund_id",3,1,15, var_refund_id))
				objCmd.Execute()				
		
				mailer_type = "customer_submitted_refund_notification"
				%>
				<!--#include virtual="emails/email_variables.asp"-->
				<%
				status = "success"
			else '======== ORDER IS NOT APPROVED 
		
				var_notes = "Automated message: Customers automated refund was declined by Authorize.net"
			
			end if '============  if response code not approved
	
		else '==============  if an error occurred

			var_notes = "Automated message: A processing error occured when customer tried to request an automated refund for the item"
			
		end if '============== if success or error message for auth.net
	

		If var_responseCode <> 1 then
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (28,?,?)"
			objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, invoice_id))
			objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,1500, var_notes))
			objCmd.Execute()
		End If

    end if ' ====== if a record is found
	

	

DataConn.Close()
Set DataConn = Nothing
%>
{"status":"<%=status%>"}