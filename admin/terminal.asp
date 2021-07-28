<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!-- #include file ="../PayPal/paypalfunctions.asp" -->
<!--#include file="../Connections/authnet.asp"-->
<html>
<head>
<link href="../CSS/Admin.css" rel="stylesheet" type="text/css" />
<title>Virtual terminal</title>
<body>
<!--#include file="admin_header.asp"-->
<br>
<br>
<div class="LargeHeader">Virtual terminal</div>
<div class="ContentText"> 
<% 
' --------- Charge Authorize.net CIM payment profile
if request.form("tender") = "charge_cim" then

		strChargeCard = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<createCustomerProfileTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <transaction>" _
		& "     <profileTransAuthCapture>" _
		& "   	   <amount>" & request.form("amount") & "</amount>" _
		& "        <customerProfileId>" & request.form("customerProfileId") & "</customerProfileId>" _
		& "   	   <customerPaymentProfileId>" & request.form("customerPaymentProfileId") & "</customerPaymentProfileId>" _
		& "   	   <customerShippingAddressId>" & request.form("customerShippingAddressId") & "</customerShippingAddressId>" _
		& "        <order>" _
		& "   	      <invoiceNumber>" & request.form("InvoiceID") & "</invoiceNumber>" _
		& "   	      <description>" & request.form("description") & "</description>" _
		& "        </order>" _
		& "   	   <taxExempt>false</taxExempt>" _
		& "   	   <recurringBilling>false</recurringBilling>" _
		& "     </profileTransAuthCapture>" _
		& "  </transaction>" _
		& "</createCustomerProfileTransactionRequest>"
	

		Set objResponseChargeCard = SendApiRequest(strChargeCard)		
		
		' APPROVED - If REGISTERED customer order is APPROVED -----------------------------------
		If IsApiResponseSuccess(objResponseChargeCard) Then
		
			strResults = objResponseChargeCard.selectSingleNode("/*/api:directResponse").Text
			response_array = split(strResults, ",")
			declineReason = response_array(3)
			
If response_array(0) = 1 Then			
%>
<div class="AccountPageHeaders" style="width: 90%;background-color:#060; color: #fff; border-color:#060;">APPROVED</div>        
<div class="AccountPageContent" style="width: 90%;background-color:#9C0; border-color:#060;">
<b>DO NOT HIT REFRESH OR BACK ... avoid duplicate charges</b>
<br />
<br />
<%= sale_type %> was successfully processed
</div> 
<%
		' ============================================
		' Write ALL info to edits log table
		' ============================================
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, invoice_id, description, edit_date) VALUES (" & user_id & "," & request.form("InvoiceID") & ",'Refunded $" & request.form("amount") & "','" & now() & "')"
		objCmd.Execute()

else ' if CIM charge was declined
%>
<div class="AccountPageHeaders" style="width: 90%;background-color:#900; color: #FFF; border-color:#900;"><%= sale_type %> DECLINED</div>        
<div class="AccountPageContent" style="width: 90%;background-color:#E9C9C9; border-color:#900;">
<strong>REASON:</strong>
<% response.write declineReason ' full text reponse from authorize.net for why declined or error %>
</div>  
<%		end if ' if response campe back from Auth.net

end if ' approved or declined

end if ' Charge CIM payment profile

' --------- Process a credit card transaction
 if request.form("tender") = "sale" OR request.form("tender") = "credit" OR request.form("tender") = "void" then
 
 ' Authorize.net get transaction details
strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
& "<getTransactionDetailsRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
& MerchantAuthentication() _
& "<transId>" & request.form("trans_id") & "</transId>" _
& "</getTransactionDetailsRequest>"

Set objGetTransactionDetails = SendApiRequest(strReq)

' If succcess retrieve transaction information
If IsApiResponseSuccess(objGetTransactionDetails) Then
	strCardNumber = objGetTransactionDetails.selectSingleNode("/*/api:transaction/api:payment/api:creditCard/api:cardNumber").Text
Else
  Response.Write "The operation failed with the following errors:<br>" & vbCrLf
  PrintErrors(objGetTransactionDetails)
End If
 
 
 If request.form("tender") = "sale" then
 	tender_type = "AUTH_CAPTURE"
	tender_email = "PAYMENT/CHARGE"
	sale_type = "PAYMENT"
 ElseIf request.form("tender") = "credit" then
 	tender_type = "CREDIT"
	tender_email = "REFUND"
	sale_type = "CREDIT"
 ElseIf request.form("tender") = "void" then
 	tender_type = "VOID"
	tender_email = "VOID"
	sale_type = "VOID"
 else
 end if
 
 ' Format cards month expiration to add a zero in front if it's a single digit month
if Len(request.form("cc_month")) = 1 then
 	cc_month = "0" & request.form("cc_month")
else
	cc_month = request.form("cc_month")
end if


Dim post_url
post_url = "https://secure.authorize.net/gateway/transact.dll"
' DEVELOPER ACCOUNTS: https://test.authorize.net/gateway/transact.dll
' REAL ACCOUNTS: https://secure.authorize.net/gateway/transact.dll

Dim post_values
Set post_values = CreateObject("Scripting.Dictionary")
post_values.CompareMode = vbTextCompare

'the API Login ID and Transaction Key must be replaced with valid values
post_values.Add "x_login", "88UD8BgvS"
post_values.Add "x_tran_key", "8842pVV38cRYyugf"
'post_values.Add "x_test_request", "TRUE" ' Turns testing on/off

post_values.Add "x_delim_data", "TRUE"
post_values.Add "x_delim_char", "|"
post_values.Add "x_relay_response", "FALSE" 'SIM applications use relay response. Set this to false if you are using AIM.

post_values.Add "x_type", tender_type 'AUTH_CAPTURE (default), AUTH_ONLY, CAPTURE_ONLY, CREDIT, PRIOR_AUTH_CAPTURE, VOID
post_values.Add "x_trans_id", request.form("trans_id") ' Reference trasaction ID # for refunds and voids

post_values.Add "x_invoice_num", request.form("InvoiceID")
post_values.Add "x_method", "CC"

if request.form("cc_num") <> "" then
	post_values.Add "x_card_num", request.form("cc_num")
	post_values.Add "x_exp_date", cc_month & request.form("cc_year")
else
	post_values.Add "x_card_num", strCardNumber
end if

post_values.Add "x_amount", request.form("amount")
post_values.Add "x_description", request.form("description") 
post_values.Add "x_email", request.form("email")
post_values.Add "x_first_name", request.form("first")
post_values.Add "x_last_name", request.form("last")
'post_values.Add "x_company", ""
'post_values.Add "x_address", ""
'post_values.Add "x_state", ""
'post_values.Add "x_zip", ""
'post_values.Add "x_country", ""
'post_values.Add "x_phone", ""

' This section takes the input fields and converts them to the proper format
' for an http post.  For example: "x_login=username&x_tran_key=a1B2c3D4"
Dim post_string 
post_string = ""
For Each Key In post_values
  post_string=post_string & Key & "=" & Server.URLEncode(post_values(Key)) & "&"
Next
post_string = Left(post_string,Len(post_string)-1)

' The following section provides an example of how to add line item details to
' the post string.  Because line items may consist of multiple values with the
' same key/name, they cannot be simply added into the above array.
'
' This section is commented out by default.
'
'Dim line_items(3)
'line_items(0) = "item1<|>golf balls<|><|>2<|>18.95<|>Y"
'line_items(1) = "item2<|>golf bag<|>Wilson golf carry bag, red<|>1<|>39.99<|>Y"
'line_items(2) = "item3<|>book<|>Golf for Dummies<|>1<|>21.99<|>Y"
'
'For Each item In line_items
'	post_string = post_string & "&x_line_item=" & Server.URLEncode(item)
'Next

' We use xmlHTTP to submit the input values and record the response
Dim objRequest, post_response
Set objRequest = Server.CreateObject("Microsoft.XMLHTTP")
'Set objRequest = Server.CreateObject("MSXML2.ServerXMLHTTP.4.0")
	objRequest.open "POST", post_url, false
	objRequest.send post_string
	post_response = objRequest.responseText
Set objRequest = nothing

' the response string is broken into an array using the specified delimiting character
Dim response_array
response_array = split(post_response, post_values("x_delim_char"), -1)

' the results are output to the screen in the form of an html numbered list.
		  'Response.Write("<OL>" & vbCrLf)
'		  For Each value in response_array
'			  Response.Write("<LI>" & value & "&nbsp;</LI>" & vbCrLf)
'		  Next
'		  Response.Write("</OL>" & vbCrLf)
' individual elements of the array could be accessed to read certain response
' fields.  For example, response_array(0) would return the Response Code,
' response_array(2) would return the Response Reason Code.
' for a list of response fields, please review the AIM Implementation Guide

If response_array(0) = 1 Then ' Result from Auth.net (1 = Approved, 2 = Declined, 3 = Error, 4 = Held for Review)

' Determing card type to send with our e-mail below
'If response_array(51) = 0  then
'	CardType = "Visa"
'ElseIf response_array(53) = 1 then
'	CardType = "Mastercard"
'ElseIf response_array(53) = 3 then
'	CardType = "American Express"
'ElseIf response_array(53) =  2 then
'	CardType = "Discover"
'ElseIf response_array(53) =  4 then
'	CardType = "Diner's Club"
'ElseIf response_array(53) =  5 then
'	CardType = "JCB"
'Else
'	CardType = "Credit card"
'End if

'Set Mail = Server.CreateObject("Bodyartforms.MailSender")
%>
<%


if request.form("clear_bo") = "yes" then

	set commUpdate = Server.CreateObject("ADODB.Command")
	commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
	commUpdate.CommandText = "UPDATE TBL_OrderSummary SET backorder = 0, notes = 'Refunded entire order $" & Request.Form("OrderTotal") & " " & date () & "' WHERE OrderDetailID = " & Request.Form("OrderDetailID") 
	commUpdate.Execute()
	Set commUpdate = Nothing
	  
	  ' Update order to cancelled/ not paid
	set UpdateOrder = Server.CreateObject("ADODB.Command")
	UpdateOrder.ActiveConnection = MM_bodyartforms_sql_STRING
	UpdateOrder.CommandText = "UPDATE sent_items SET shipped = 'Cancelled', ship_code = 'not paid', USPS_tracking = '', UPS_tracking = ''  WHERE ID = " & Request.Form("InvoiceID") & "" 
	UpdateOrder.Execute()

end if ' clear off backorder status

if request.form("clear_bo_item") = "yes" then ' clear backorder just for ONE item

	set commUpdate = Server.CreateObject("ADODB.Command")
	commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
	commUpdate.CommandText = "UPDATE TBL_OrderSummary SET backorder = 0, item_price = 0, notes = 'Out; Refunded $" & Request.Form("amount") & " " & date () & "' WHERE OrderDetailID = " & Request.Form("OrderDetailID") 
	commUpdate.Execute()
	Set commUpdate = Nothing

end if
%>
<div class="AccountPageHeaders" style="width: 90%;background-color:#060; color: #fff; border-color:#060;">APPROVED</div>        
<div class="AccountPageContent" style="width: 90%;background-color:#9C0; border-color:#060;">
<b>DO NOT HIT REFRESH OR BACK ... avoid duplicate charges</b>
<br />
<br />
<%= sale_type %> was successfully processed
</div>                   
<% 
			' ============================================
			' Write ALL info to edits log table
			' ============================================
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, invoice_id, description, edit_date) VALUES (" & user_id & "," & request.form("InvoiceID") & ",'Refunded $" & request.form("amount") & "','" & now() & "')"
			objCmd.Execute()


Else' If the credit card is declined
%>
<div class="AccountPageHeaders" style="width: 90%;background-color:#900; color: #FFF; border-color:#900;"><%= sale_type %> DECLINED</div>        
<div class="AccountPageContent" style="width: 90%;background-color:#E9C9C9; border-color:#900;">
<strong>REASON:</strong>
<% response.write response_array(3) ' full text reponse from authorize.net for why declined or error %>
</div>   

<%
End If ' If the order is approved
end if ' being processed with a credit card %>  
<% If request.form("tender") = "" then %>


 <form action="terminal.asp" method="post">
   <p>
     <input type="radio" name="tender" value="sale" id="type_0" checked="checked">
     Sale&nbsp;&nbsp;&nbsp;
  <input type="radio" name="tender" value="credit" id="type_0">
    Refund</p>
   <p>Transaction ID (if any):<br>
     <input name="trans_id" type="text" id="trans_id" size="30" style="margin: 2px;" />
   </p>
   <p>Invoice # (if any) 
     <input name="invoice" type="text" id="invoice" style="margin: 2px;" size="10" maxlength="10" />
   </p>
   <p>Amount: 
     &nbsp;&nbsp;&nbsp;$
<input name="amount" type="text" id="amount" style="margin: 2px;" size="10" maxlength="10" />
   </p>
   <p>Card #
     <input name="cc_num" type="text" id="cc_num" style="margin: 2px;" size="20" maxlength="20" />
   &nbsp;&nbsp;&nbsp;Exp: 
     <input name="cc_month" type="text" id="cc_month" style="margin: 2px;" size="2" maxlength="2" /> 
     / 
     <input name="cc_year" type="text" id="cc_year" style="margin: 2px;" size="2" maxlength="2" />
   </p>
   <p>First name:
     <input name="first" type="text" id="first" size="20" style="margin: 2px;" />
   Last name: 
   <input name="last" type="text" id="last" size="20" style="margin: 2px;" />
   </p>
   <p>E-mail:
     <input name="email" type="text" id="email" size="30" style="margin: 2px;" />
   </p>
   <p>Description (will be sent to customer):<br>
     <textarea name="description" cols="50" rows="4" id="description" style="margin: 2px;"></textarea>
     <br>
   </p>
   <p>
     <input type="submit" value="Submit">
 </p>
 </form>
 <% else %><br>
<br>

 <a href="terminal.asp">Click here to run a new transaction</a>
 <% if request.form("InvoiceID") <> "" then %>
 <br/><br/>
 <a href="invoice.asp?ID=<%= request.form("InvoiceID") %>">Click here to return to invoice</a>
 <% end if %>
 <% end if %>
</div><br>

</body>
</html>
