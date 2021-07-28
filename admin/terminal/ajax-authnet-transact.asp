<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/authnet.asp" -->
<%
' only run through auth.net if it's not a paypal money request
if request.form("tender") <> "money-request" then

if request.form("tender") = "charge_cim" then
	sale_type = "authCaptureTransaction"
	transaction_id = ""
	str_cim = "<profile><customerProfileId>" & request.form("customerProfileId") & "</customerProfileId><paymentProfile><paymentProfileId>" & request.form("customerPaymentProfileId") & "</paymentProfileId></paymentProfile></profile>"
	var_amount = "<amount>" & request.form("amount") & "</amount>"
elseif  request.form("tender") = "refund" then
	sale_type = "refundTransaction"
	transaction_id = "<refTransId>" & request.form("trans_id") & "</refTransId>"
	str_cim = ""
	var_amount = "<amount>" & request.form("amount") & "</amount>"
	if request.form("pay-method") <> "PayPal" then
		var_card_info = "<payment><creditCard><cardNumber>" & request.form("card_number") & "</cardNumber><expirationDate>XXXX</expirationDate></creditCard></payment>"
	else
		var_card_info = ""
	end if
elseif  request.form("tender") = "void" then
	sale_type = "voidTransaction"
	transaction_id = "<refTransId>" & request.form("trans_id") & "</refTransId>"
	str_cim = ""
	var_amount = ""
end if


strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" _
& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
& MerchantAuthentication() _
& "<transactionRequest>" _
& "		<transactionType>" & sale_type & "</transactionType>" _
		& var_amount _
		& var_card_info _
		& str_cim _
		& transaction_id _
& "		<order>" _
& "			<invoiceNumber>" & request.form("invoice") & "</invoiceNumber>" _
& "			<description>" & request.form("description") & "</description>" _
& "		</order>" _
& "</transactionRequest>" _
& "</createTransactionRequest>"

Set objResponse = SendApiRequest(strSend)

	var_message = objResponse.selectSingleNode("/*/api:messages/api:message/api:text").Text

' APPROVED - If REGISTERED customer order is APPROVED -----------------------------------
If IsApiResponseSuccess(objResponse) Then

	var_responseCode = objResponse.selectSingleNode("/*/api:transactionResponse/api:responseCode").Text

	if var_responseCode = 1 then ' approved 
%>
	{  
		"status":"success",
		"reason":"<%= var_message %>"
	}
<%	

	' ============================================
	' Write ALL info to edits log table
	' ============================================
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, invoice_id, description, edit_date) VALUES (" & user_id & "," & request.form("invoice") & ",'Refunded $" & request.form("amount") & "','" & now() & "')"
	objCmd.Execute()

else ' if not approved 
	if var_responseCode = 2 then
		var_message = "Declined"
	elseif  var_responseCode = 3 then
		var_message = "Error"
	else
		var_message = "Held for review"
	end if

%>
	{  
		"status":"decline",
		"reason":"<%= var_message %>"
	}
<%	end if ' if response code not approved

else ' if an error occurred
%>
	{  
		"status":"decline",
		"reason":"<%= var_message %>"
	}
 
<%	end if ' if success or error message for auth.net 



' run if it's a paypal money request
else	'  request.form("tender") = "money-request"
	mailer_type = "money-request"
%>
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/emails/email_variables.asp"-->
	{  
		"status":"success",
		"reason":"PayPal money request email has been sent"
	}
<%	

end if 	'  request.form("tender") = "money-request"


%>