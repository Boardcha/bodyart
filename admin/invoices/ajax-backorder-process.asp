<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/authnet.asp" -->

<%
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT customer_ID, total_preferred_discount, total_coupon_discount, total_sales_tax, transactionID, pay_method, email, customer_first, ID FROM sent_items WHERE ID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("string_id",3,1,12,request.form("invoice")))
Set rsGetOrder = objCmd.Execute()

var_custid = 0
var_coupon_discount = 0
var_preferred_discount = 0
var_sales_tax = 0
if not rsGetOrder.eof then
	var_custid = rsGetOrder.Fields.Item("customer_ID").Value
	var_coupon_discount = rsGetOrder.Fields.Item("total_coupon_discount").Value
	var_preferred_discount = rsGetOrder.Fields.Item("total_preferred_discount").Value
	var_sales_tax = rsGetOrder.Fields.Item("total_sales_tax").Value
	transaction_id = rsGetOrder.Fields.Item("transactionID").Value
	pay_method = rsGetOrder.Fields.Item("pay_method").Value
	
	if pay_method <> "PayPal" then
		var_card_info = "<payment><creditCard><cardNumber>" & request.form("card_number") & "</cardNumber><expirationDate>XXXX</expirationDate></creditCard></payment>"
	else
		var_card_info = ""
	end if
	
end if

'RESPONSE.WRITE "XML STRING: " & var_card_info

var_invoice = request.form("invoice")
var_item = request.form("item")
var_qty = request.form("qty")
var_detailid = request.form("detailid")
var_price = request.form("price")
var_origprice = request.form("origprice")
var_total = request.form("total")
var_exchange_detailid = request.form("exchange_detailid")
var_exchange_productid = request.form("exchange_productid")
var_exchange_qty = request.form("exchange_qty")
var_exchange_origitem = request.form("exchange_origitem")
var_exchange_newitem = ""
var_exchange_price_diff = request.form("exchange_price_diff")

' Retrieve the current order
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM sent_items WHERE ID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,12,var_invoice))
Set rsGetOrder = objCmd.Execute()

if rsGetOrder("country") <> "USA" AND rsGetOrder("country") <> "US" then
	shipping = "DHL GlobalMail Packet Priority"
else
	shipping = "DHL Basic mail"
end if


' Retrieve product ID to get title
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT *, ProductDetailID, title + ' ' + Gauge + ' ' + Length + ' ' + ProductDetail1 AS 'item_title' FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID WHERE ProductDetailID = ?" 
objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,12, var_detailid))
Set rsGetItemDetails = objCmd.Execute()


function NewOrder
	
	set NewOrder = Server.CreateObject("ADODB.Command")
	NewOrder.ActiveConnection = DataConn
	NewOrder.CommandText = "INSERT INTO sent_items (shipped, customer_ID, customer_first, customer_last, company, address, address2, city, state, province, zip, country, email, date_order_placed, shipping_rate, shipping_type, item_description, ship_code, phone, UPS_Service) SELECT 'Pending...', customer_ID, customer_first, customer_last, company, address, address2, city, state, province, zip, country, email, '" & now() & "', 0,'" & shipping & "','<b><font size=3>BACKORDER SHIPMENT</font></B><br>', 'paid', phone, '' FROM sent_items WHERE ID =" & var_invoice 
	NewOrder.Execute() 
	
	' Retrieve the newest order
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT ID, email FROM dbo.sent_items WHERE email = '" & rsGetOrder.Fields.Item("email").Value & "' ORDER BY ID DESC" 
	Set NewOrder = objCmd.Execute()

end function ' new order

function DeductQuantities

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE ProductDetails SET qty = qty - " & var_qty & ", DateLastPurchased = '"& date() &"' WHERE ProductDetailID = " & var_detailid 
	objCmd.Execute()

end function ' Deduct Quantities

function CancelOrder
	
	' Update order to cancelled/ not paid
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET shipped = 'Cancelled', ship_code = 'not paid' WHERE ID = " & var_invoice
	objCmd.Execute()
	
end function ' 	Cancel Order

function CalcCoupons

	' If there was a coupon used then reduce the coupon and  tax amount on the order
	if var_price < var_origprice OR var_sales_tax > 0 then
		var_discount_difference = var_origprice - var_price
		item_tax = FormatNumber(var_price * .0825, -1, -2, -0, -2)
		
			discount = ""
		if var_coupon_discount <> 0 then
			discount = "total_coupon_discount = total_coupon_discount - " & var_discount_difference
			response.write "test"
		end if
		if var_preferred_discount <> 0 then
			discount = "total_preferred_discount = total_preferred_discount - " & var_discount_difference
		end if
		if var_coupon_discount <> 0 or var_preferred_discount <> 0 then
			add_comma = ", "
		end if
		if var_sales_tax > 0 then
			write_tax = "total_sales_tax = total_sales_tax - " & item_tax
		else
			write_tax = "total_sales_tax = 0 "
		end if

		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE sent_items SET " & discount & " " & add_comma & " " & write_tax & " WHERE ID = ?" 
		objCmd.Parameters.Append(objCmd.CreateParameter("string_id",3,1,12,var_invoice))
		objCmd.Execute()

	end if	'	var_price < var_origprice 

end function	'	CalcCoupons

function SetQtyZero

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_OrderSummary SET item_price = 0 WHERE OrderDetailID = " & var_item 
	objCmd.Execute()
	
end function	'	SetQtyZero


'	Reship just the backorder item ------------------------
if request.form("agenda") = "ship-one" then
 
	set rsGetNewInvoice = NewOrder()

	set CopyItem = Server.CreateObject("ADODB.Command")
	CopyItem.ActiveConnection = DataConn
	CopyItem.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, ProductID, DetailID, qty) SELECT " & rsGetNewInvoice.Fields.Item("ID").Value & ", ProductID, DetailID, qty FROM TBL_OrderSummary WHERE OrderDetailID =" & var_item
	CopyItem.Execute() 
	
	DeductQuantities()
	
	var_notes = "Automated message: Backordered item # " & var_detailid & " came back in stock. Created new order to ship out."
	new_invoice_id = rsGetNewInvoice.Fields.Item("ID").Value

	mailer_type = "bo_notification"
	backorder_email_body = "The item is now back in stock and we have set up a new order to ship it out (New order # " & new_invoice_id & "). You will be e-mailed with a separate shipping notification along with a tracking # once the order ships."

%>
{
	"status":"Backorder shipment created</br><strong><ul><li>CUSTOMER HAS BEEN EMAILED</li><li>Quantities have been deducted</li></ul></strong><a href='invoice.asp?ID=<%= rsGetNewInvoice.Fields.Item("ID").Value %>'>Click here</a> to go to new order"
}
<%
end if '	Reship just the backorder item 


' 	Reship the entire order ----------------------------------
if request.form("agenda") = "reship" then
 
	set CopyItem = Server.CreateObject("ADODB.Command")
	CopyItem.ActiveConnection = DataConn
	CopyItem.CommandText = "UPDATE sent_items SET shipped = 'Pending...' WHERE ID = " & var_invoice
	CopyItem.Execute() 
	
	DeductQuantities()
	
	var_notes = "Automated message: Backordered item # " & var_detailid & " came back in stock. Set entire order to reship."
	mailer_type = "bo_notification"
	backorder_email_body = "The item is now back in stock and your order will be shipping out. You will be e-mailed with a separate shipping notification along with a tracking # once the order ships."

%>
{
	"status":"Current order has been set to ship out again.</br><strong><ul><li>CUSTOMER HAS BEEN EMAILED</li><li>Quantities have been deducted</li></ul></strong><a href='invoice.asp?ID=<%= var_invoice %>'>Click here</a> to refresh order"
}
<%
end if '	Reship the entire order


'	Take item off backorder status ------------------------------
if request.form("agenda") = "clear" then
 
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_OrderSummary SET backorder = 0 WHERE OrderDetailID =" & var_item 
	objCmd.Execute()
	
	var_notes = "Automated message: Cleared backorder item # " & var_detailid

	if request.form("stock_qty") <> "" then
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE ProductDetails SET qty = ? WHERE ProductDetailID = ?" 
		objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,12, request.form("stock_qty") ))
		objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,12, var_detailid )) 
		objCmd.Execute()
	end if 

%>
{
	"status":"Backorder has been cleared</br><br/><a href='invoice.asp?ID=<%= var_invoice %>'>Click here</a> to refresh order"
}
<%
end if '	Take item off backorder status

' 	Issue a store credit for the ITEM ONLY -----------------------------
if request.form("agenda") = "item-storecredit" then
 
	' Update customer credits 
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET credits = credits + " & var_price & " WHERE customer_ID = " & var_custid 
	objCmd.Execute()
	
	CalcCoupons()
	SetQtyZero()
	
	var_notes = "Automated message: Customer opted for a store credit $" & var_price & " on backordered item # " & var_detailid
	mailer_type = "bo_notification"
	backorder_email_body = "Per your request we have issued a store credit in the amount of $" & var_price & " to your account and the funds are available for use immediately."

%>
{
	"status":"Store credit has been issued.</br><strong><ul><li>CUSTOMER HAS BEEN EMAILED</li><li>Coupon & tax (if applicable) have been adusted.</li></strong><a href='invoice.asp?ID=<%= var_invoice %>'>Click here</a> to refresh order"
}
<%
end if ' 	Issue a store credit for the ITEM ONLY


' 	Issue a store credit for the ENTIRE ORDER ---------------------------
if request.form("agenda") = "cancel-storecredit" then
 
	CancelOrder()
	
	' Update customer credits 
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET credits = credits + " & var_total & " WHERE customer_ID = " & var_custid 
	objCmd.Execute()
	
	var_notes = "Automated message: Customer opted to cancel the entire order for a store credit $" & var_total & " on backordered item # " & var_detailid
	mailer_type = "bo_notification"
	backorder_email_body = "Per your request we have cancelled the order and issued a store credit in the amount of $" & var_total & " to your account and the funds are available for use immediately."

%>
{
	"status":"Store credit has been issued.<strong><ul><li>CUSTOMER HAS BEEN EMAILED</li></strong><a href='invoice.asp?ID=<%= var_invoice %>'>Click here</a> to refresh order"
}
<%
end if ' 	Issue a store credit for the ENTIRE ORDER


' 	Issue a REFUND for the ITEM ONLY ---------------------------
if request.form("agenda") = "item-refund" then
 
	strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" _
	& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
	& MerchantAuthentication() _
	& "<transactionRequest>" _
	& "		<transactionType>refundTransaction</transactionType>" _
	& "		<amount>" & var_price & "</amount>" _
			& var_card_info _
	& "		<refTransId>" & transaction_id & "</refTransId>" _
	& "		<order>" _
	& "			<invoiceNumber>" & var_invoice & "</invoiceNumber>" _
	& "			<description>Backorder refund</description>" _
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
			"status":"<ul><li><strong>CUSTOMER HAS BEEN E-MAILED</strong></li><li>Refund has been issued</li></ul><a href='invoice.asp?ID=<%= var_invoice %>'>Click here</a> to refresh order"
		}
	<%	
		
		CalcCoupons()
		SetQtyZero()
	
		var_notes = "Automated message: Customer opted to refund item only $" & var_price & " on backordered item # " & var_detailid
		mailer_type = "bo_notification"
		backorder_email_body = "Per your request we have cancelled the item and issued a refund to your credit card in the amount of $" & var_price & ". This refund should be back on your card in about 5-7 business days."
	
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
			"status":"<div class='notice-red'>DECLINED, <%= var_message %></div>"
		}
	<%	
		var_notes = "Automated message: Declined refund for item only $" & var_price & " on backordered item # " & var_detailid
	
		end if ' if response code not approved

	else ' if an error occurred
	%>
		{  
			"status":"<div class='notice-red'>ERROR PROCESSING REQUEST - <%= var_message %> <%= objResponse.selectSingleNode("/*/api:transactionResponse/api:errors/api:error/api:errorText").Text %></div>"
		}
	 
	<%	
		var_notes = "Automated message: ERROR Processing refund for item only $" & var_price & " on backordered item # " & var_detailid
		
		end if ' if success or error message for auth.net

end if ' 	Issue a REFUND for the ITEM ONLY



' 	Issue a REFUND for the ENTIRE ORDER ---------------------------
if request.form("agenda") = "cancel-refund" then
 
	strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" _
	& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
	& MerchantAuthentication() _
	& "<transactionRequest>" _
	& "		<transactionType>refundTransaction</transactionType>" _
	& "		<amount>" & var_total & "</amount>" _
			& var_card_info _
	& "		<refTransId>" & transaction_id & "</refTransId>" _
	& "		<order>" _
	& "			<invoiceNumber>" & var_invoice & "</invoiceNumber>" _
	& "			<description>Backorder refund</description>" _
	& "		</order>" _
	& "</transactionRequest>" _
	& "</createTransactionRequest>"
	
	Set objResponse = SendApiRequest(strSend)

		var_message = objResponse.selectSingleNode("/*/api:messages/api:message/api:text").Text

	' APPROVED - If REGISTERED customer order is APPROVED -----------------------------------
	If IsApiResponseSuccess(objResponse) Then

		var_responseCode = objResponse.selectSingleNode("/*/api:transactionResponse/api:responseCode").Text

		if var_responseCode = 1 then ' approved 
		
			CancelOrder()
	%>
		{  
			"status":"<ul><li><strong>CUSTOMER HAS BEEN E-MAILED</strong></li><li>Refund has been issued</li></ul><a href='invoice.asp?ID=<%= var_invoice %>'>Click here</a> to refresh order"
		}
	<%	
	
		var_notes = "Automated message: Customer opted to cancel the entire order for a refund $" & var_total & " on backordered item # " & var_detailid
		mailer_type = "bo_notification"
		backorder_email_body = "Per your request we have cancelled the entire order and issued a refund to your credit card in the amount of $" & var_total & ". This refund should be back on your card in about 5-7 business days."
	
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
			"status":"<div class='notice-red'>DECLINED, <%= var_message %></div>"
		}
	<%	
		var_notes = "Automated message: Declined refund for entire order $" & var_total & " on backordered item # " & var_detailid
	
		end if ' if response code not approved

	else ' if an error occurred
	%>
		{  
			"status":"<div class='notice-red'>ERROR PROCESSING REQUEST - <%= var_message %> <%= objResponse.selectSingleNode("/*/api:transactionResponse/api:errors/api:error/api:errorText").Text %></div>"
		}
	 
	<%	
		var_notes = "Automated message: ERROR Processing refund for entire order $" & var_total & " on backordered item # " & var_detailid
		
		end if ' if success or error message for auth.net

end if ' 	Issue a REFUND for the ENTIRE ORDER


'	Exchange item -----------------------------------------
if request.form("agenda") = "exchange" then

	'====== GET INFORMATION ABOUT NEW EXCHANGED ITEM ======
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT *, ProductDetailID, title + ' ' + Gauge + ' ' + Length + ' ' + ProductDetail1 AS 'exchange_item_title' FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID WHERE ProductDetailID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,12, var_exchange_detailid))
	Set rsGetExchangeItemDetails = objCmd.Execute()
 
	set rsGetNewInvoice = NewOrder()

	function processExchange()
		if request.form("exchange_agenda") = "amount-owed" then
			var_add_price = var_exchange_price_diff
		else
			var_add_price = 0
		end if
	
		' Add exchanged item into new order
		set AddItem = Server.CreateObject("ADODB.Command")
		AddItem.ActiveConnection = DataConn
		AddItem.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, ProductID, DetailID, qty, item_price) VALUES (" & rsGetNewInvoice.Fields.Item("ID").Value & ", " & var_exchange_productid & ", " & var_exchange_detailid & ", " & var_exchange_qty & ", " & var_add_price & ")"
		AddItem.Execute() 
		
		' DEDUCT QUANTITIES FOR DIFFERENT PRODUCT
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE ProductDetails SET qty = qty - " & var_exchange_qty & ", DateLastPurchased = '"& date() &"' WHERE ProductDetailID = " & var_exchange_detailid 
		objCmd.Execute()
	
	end function	'	processExchange()
	
	
	' Exchange with store credit being due
	if request.form("exchange_agenda") = "storecredit" then
	
		processExchange()
		
		' Update customer credits 
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE customers SET credits = credits + " & var_exchange_price_diff & " WHERE customer_ID = " & var_custid 
		objCmd.Execute()
		
		SetQtyZero()
		
		var_notes = "Automated message: Customer exchanged backordered item #" & var_exchange_origitem & " " & rsGetItemDetails.Fields.Item("item_title").Value & " for item on invoice #" & rsGetNewInvoice.Fields.Item("ID").Value & ". Price difference of $" & var_exchange_price_diff & " has been applied to store credit. Created new order to ship out."
		mailer_type = "bo_notification"
		backorder_email_body = "Per your request we have exchanged the item for " & rsGetExchangeItemDetails.Fields.Item("exchange_item_title").Value & " (new invoice #" & rsGetNewInvoice.Fields.Item("ID").Value & "). The price difference of $" & var_exchange_price_diff & " has been applied as a store credit on your account and the funds are available for use immediately.<br/><br/>Your exchanged item will ship out soon and you will be e-mailed with a separate shipping notification along with a tracking #."
		new_invoice_id = rsGetNewInvoice.Fields.Item("ID").Value

		%>
		{
			"status":"<ul><li><strong>CUSTOMER HAS BEEN EMAILED</strong></li><li>Backorder shipment created</li><li>Quantities have been deducted for exchanged item.<li>Store credit has been issued in the amount of $<%= request.form("exchange_price_diff") %></li></ul><a href='invoice.asp?ID=<%= rsGetNewInvoice.Fields.Item("ID").Value %>'>Click here</a> to go to new order"
		}
		<%
		
	' Exchange and issue a credit card refund for difference
	elseif request.form("exchange_agenda") = "cardrefund" then
		
		' Exchange and issue a credit card refund for difference
		strSend = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "<transactionRequest>" _
		& "		<transactionType>refundTransaction</transactionType>" _
		& "		<amount>" & var_exchange_price_diff & "</amount>" _
				& var_card_info _
		& "		<refTransId>" & transaction_id & "</refTransId>" _
		& "		<order>" _
		& "			<invoiceNumber>" & var_invoice & "</invoiceNumber>" _
		& "			<description>Backorder refund</description>" _
		& "		</order>" _
		& "</transactionRequest>" _
		& "</createTransactionRequest>"
		
		Set objResponse = SendApiRequest(strSend)

			var_message = objResponse.selectSingleNode("/*/api:messages/api:message/api:text").Text

		If IsApiResponseSuccess(objResponse) Then

			var_responseCode = objResponse.selectSingleNode("/*/api:transactionResponse/api:responseCode").Text

		' Exchange and issue a credit card refund for difference
		
		' APPROVED    -----------------------------------
			if var_responseCode = 1 then ' approved 
			
			processExchange()
		%>
		{
			"status":"<ul><li><strong>CUSTOMER HAS BEEN EMAILED</strong></li><li>Backorder shipment created</li><li>Quantities have been deducted for exchanged item.</li><li>Auth.net credit has been issued in the amount of $<%= request.form("exchange_price_diff") %></li></ul><a href='invoice.asp?ID=<%= rsGetNewInvoice.Fields.Item("ID").Value %>'>Click here</a> to go to new order"
		}
		<%	
			' Exchange and issue a credit card refund for difference
			' Update customer credits 
			set objCmd = Server.CreateObject("ADODB.Command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "UPDATE customers SET credits = credits + " & var_exchange_price_diff & " WHERE customer_ID = " & var_custid 
			objCmd.Execute()
			
			SetQtyZero()
			
			var_notes = "Automated message: Customer exchanged backordered item #" & var_exchange_origitem & " for item on invoice #" & rsGetNewInvoice.Fields.Item("ID").Value & ". Price difference of $" & var_exchange_price_diff & " has been refunded via Auth.net. Created new order to ship out."
			mailer_type = "bo_notification"
			backorder_email_body = "Per your request we have exchanged the item for " & rsGetExchangeItemDetails.Fields.Item("exchange_item_title").Value & " (new invoice #" & rsGetNewInvoice.Fields.Item("ID").Value & "). The price difference of $" & var_exchange_price_diff & " has been refunded to your card and you should see the funds available in 5-7 business days.<br/><br/>Your exchanged item will ship out soon and you will be e-mailed with a separate shipping notification along with a tracking #."
			new_invoice_id = rsGetNewInvoice.Fields.Item("ID").Value
		
		' Exchange and issue a credit card refund for difference
		'	DECLINED  ---------------------------------
		
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
				"status":"<div class='notice-red'>DECLINED, <%= var_message %></div>"
			}
		<%	
			var_notes = "Automated message: Declined exchange refund due for $" & var_exchange_price_diff & " on backordered item # " & var_exchange_origitem
		
			end if ' if response code not approved
		
		' Exchange and issue a credit card refund for difference
		'	ERROR  ---------------------------------
		else ' if an error occurred
		%>
			{  
				"status":"<div class='notice-red'>ERROR PROCESSING REQUEST - <%= var_message %> <%= objResponse.selectSingleNode("/*/api:transactionResponse/api:errors/api:error/api:errorText").Text %></div>"
			}
		 
		<%	
			var_notes = "Automated message: ERROR Processing exchange refund due for $" & var_exchange_price_diff & " on backordered item # " & var_exchange_origitem
			
		end if ' Exchange and issue a credit card refund for difference
	
	
	elseif request.form("exchange_agenda") = "equal-exchange" then '	For anything that is an equal exchange $0 value

		processExchange()
		
		var_notes = "Automated message: Customer exchanged backordered item #" & var_exchange_origitem & " for item on invoice #" & rsGetNewInvoice.Fields.Item("ID").Value & ". Exchange was of equal value. Created new order to ship out."
		mailer_type = "bo_notification"
		backorder_email_body = "Per your request we have exchanged the item for " & rsGetExchangeItemDetails.Fields.Item("exchange_item_title").Value & " (new invoice #" & rsGetNewInvoice.Fields.Item("ID").Value & "). Your exchanged item will ship out soon and you will receive an e-mail with a shipping notification along with a tracking #."
		new_invoice_id = rsGetNewInvoice.Fields.Item("ID").Value

		%>
		{
			"status":"Backorder shipment created</br><strong><ul><li>CUSTOMER HAS BEEN EMAILED</li><li>Quantities have been deducted for exchanged item.</li><li>Exchange was of equal value.</li></ul></strong><a href='invoice.asp?ID=<%= rsGetNewInvoice.Fields.Item("ID").Value %>'>Click here</a> to go to new order"
		}
		<%

			
		elseif request.form("exchange_agenda") = "amount-owed" then ' For an exchange where there is an amount due/owed by customer

		processExchange()

		var_notes = "Automated message: Customer exchanged backordered item #" & var_exchange_origitem & " for item on invoice #" & rsGetNewInvoice.Fields.Item("ID").Value & ". Amount owed is $" & var_exchange_price_diff & ". Created new order to ship out."
		new_invoice_id = rsGetNewInvoice.Fields.Item("ID").Value

		%>
		{
		"status":"<ul><li><strong>YOU WILL NEED TO E-MAIL CUSTOMER WITH PAYMENT INFORMATION FOR AMOUNT DUE</strong></li><li>Backorder shipment created</li><li>Quantities have been deducted for exchanged item.</li></ul><a href='invoice.asp?ID=<%= rsGetNewInvoice.Fields.Item("ID").Value %>'>Click here</a> to go to new order"
		}
		<%

	else
		'====== do nothing
	end if 	'	Exchange type (store credit, card refund, or equal exchange)
	
end if '	End exchange item  ------------------------------------- 




' 	============== DO TASKS BELOW FOR EACH SCENARIO ======================
' Notes for original order
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?,?)"
objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10,user_id))
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10,var_invoice))
objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,250,var_notes))
objCmd.Execute()

' Notes for NEW order created
if new_invoice_id <> "" then
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?,?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10,user_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10,new_invoice_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,250,"Automated message: Backorder shipment generated from another order"))
	objCmd.Execute()
end if


' Clear backorder status from any item that accesses this page
	set clearbo = Server.CreateObject("ADODB.Command")
	clearbo.ActiveConnection = MM_bodyartforms_sql_STRING
	clearbo.CommandText = "UPDATE TBL_OrderSummary SET backorder = 0, BackorderReview = 'N', notes = 'Removed BO status' WHERE OrderDetailID =" & var_item 
	clearbo.Execute()

	%>
	<!--#include virtual="emails/function-send-email.asp"-->
	<!--#include virtual="/emails/email_variables.asp"-->
	<%

DataConn.Close()
%>
