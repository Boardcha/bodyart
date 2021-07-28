<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/authnet.asp" -->
<!--#include virtual="emails/function-send-email.asp"-->
{
<%
' ============== UNFINISHED TASKS AS OF SEPT 2019 ====================
' What about orders that have add-on items from two different transactions
' Order with autoclave service... autoclave service needs to be refunded if all autoclavable items are returned?
' ============== UNFINISHED TASKS AS OF SEPT 2019 ====================

' ====== set variables
var_invoiceid = request.form("invoice")
return_extra_comments = request.form("return-extra-comments")
var_gift_cert_code_used = request.form("var_gift_cert_code_used")
storecredit_refund_due = CCur(request.form("returns-storecredit-due"))
giftcert_refund_due = CCur(request.form("returns-giftcert-due"))
cc_refund_due = CCur(request.form("returns-ccrefund"))
returns_sales_tax = CCur(request.form("returns-sales-tax"))
returns_card_number = request.form("returns_card_number")
returns_calculation = request.form("returns-calculation")

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM sent_items WHERE ID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("string_id",3,1,12,var_invoiceid))
Set rsGetOrder = objCmd.Execute()

Set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM TBL_OrderSummary WHERE InvoiceID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, var_invoiceid))
Set rsGetOrderItems = objCmd.Execute()

if not rsGetOrder.eof then
	var_custid = rsGetOrder.Fields.Item("customer_ID").Value
	transaction_id = rsGetOrder.Fields.Item("transactionID").Value
	pay_method = rsGetOrder.Fields.Item("pay_method").Value
	
	if pay_method <> "PayPal" then
		var_card_info = "<payment><creditCard><cardNumber>" & returns_card_number & "</cardNumber><expirationDate>XXXX</expirationDate></creditCard></payment>"
	else
		var_card_info = ""
	end if
	
end if

' =============== Issue STORE CREDIT for returns ===============
if storecredit_refund_due > 0 then
 
	' Update customer credits 
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET credits = credits + " & storecredit_refund_due & " WHERE customer_ID = " & var_custid 
	objCmd.Execute()
	%>
			"status":"success",
			"status_store_credit":"success",
	<%
end if 

' =============== Issue GIFT CERTIFICATE for returns ===============
if giftcert_refund_due > 0 then
 
	' Update customer credits 
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBLcredits SET amount = amount + " & giftcert_refund_due & " WHERE code = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("code",200,1,150,var_gift_cert_code_used))
	objCmd.Execute()

	%>
			"status":"success",
			"status_gift_cert":"success",
	<%

end if 


' =============== Issue REFUND for returns ===============
if cc_refund_due > 0 then
 
	strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" _
	& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
	& MerchantAuthentication() _
	& "<transactionRequest>" _
	& "		<transactionType>refundTransaction</transactionType>" _
	& "		<amount>" & cc_refund_due & "</amount>" _
			& var_card_info _
	& "		<refTransId>" & transaction_id & "</refTransId>" _
	& "		<order>" _
	& "			<invoiceNumber>" & var_invoiceid & "</invoiceNumber>" _
	& "			<description>Refund for returned items</description>" _
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
			"status":"success",
	<%	
	else ' if not approved 
			if var_responseCode = 2 then
				var_message = "Declined"
			elseif  var_responseCode = 3 then
				var_message = "Error"
			else
				var_message = "Held for review"
			end if

	%>
			"status":"<div class='notice-red'>DECLINED, <%= var_message %></div>",
	<%	
		end if ' if response code not approved

	else ' if an error occurred
	%>
			"status":"<div class='notice-red'>ERROR PROCESSING REQUEST - <%= var_message %> <%= objResponse.selectSingleNode("/*/api:transactionResponse/api:errors/api:error/api:errorText").Text %></div>",
	 
	<%		
		end if ' if success or error message for auth.net


end if ' ======== Issue a CARD or PAYPAL refund



' 	============== DO TASKS BELOW FOR EACH SCENARIO ======================
While NOT rsGetOrderItems.EOF 
For Each item In Request.Form
If IsNumeric(item) Then
	' --- if form name is integer and matches then write item notes 
	If Clng(item) = CLng(rsGetOrderItems.Fields.Item("OrderDetailID").Value) Then
		Set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE TBL_OrderSummary SET returned = 1, return_date = GETDATE(), returned_qty = ? WHERE OrderDetailID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,8, request.form(item)))
		objCmd.Parameters.Append(objCmd.CreateParameter("OrderDetailID",3,1,20, Clng(item)))
		objCmd.Execute()

		Set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT d.active AS d_active, p.active AS p_active, p.title FROM ProductDetails AS d INNER JOIN jewelry AS p ON d.ProductID = p.ProductID WHERE d.ProductDetailID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,20, rsGetOrderItems.Fields.Item("DetailID").Value))
		set rsCheckStockLevels = objCmd.Execute()

		' ===== put item back in stock
			Set objCmd = Server.CreateObject("ADODB.Command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "UPDATE ProductDetails SET qty =  qty + ? WHERE ProductDetailID = ?"
			objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,8, request.form(item)))
			objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,20, rsGetOrderItems.Fields.Item("DetailID").Value))
			objCmd.Execute()

			if rsCheckStockLevels.Fields.Item("d_active").Value = 0 or rsCheckStockLevels.Fields.Item("p_active").Value = 0 then
				var_active_status = var_active_status & " " & rsCheckStockLevels.Fields.Item("title").Value & "<br/>"
			end if 
		set rsCheckStockLevels = nothing
	end if
end if
Next
rsGetOrderItems.MoveNext()
Wend

%>
		"var_active_status":"<%= var_active_status %>"
<%
' Notes for original order
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?,?)"
objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10,user_id))
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,var_invoiceid))
objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,1500, return_extra_comments & "<br/><br/>" & returns_calculation))
objCmd.Execute()


' Update total returns in DB and tax returns portion 
set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE sent_items SET total_returns = total_returns + " & cc_refund_due + storecredit_refund_due + giftcert_refund_due  & ", returns_sales_tax = returns_sales_tax + " & returns_sales_tax & " WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,var_invoiceid))

objCmd.Execute()

mailer_type = "returns"
var_reason = "return_refunded"
%>
<!--#include virtual="emails/email_variables.asp"-->
<%

DataConn.Close()
%>
	}