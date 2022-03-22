<%@LANGUAGE="VBSCRIPT"%>
<%
Server.ScriptTimeout = 240
%>
<%
	page_title = "Bodyartforms PayPal checkout"
	page_description = "Bodyartforms paypal checkout"
	page_keywords = "body jewelry, shopping cart, basket"

	If IsNumeric(session("third_party_total")) Then
		if session("third_party_total") > 0 then
			paypal_amt = FormatNumber(session("third_party_total"), -1, -2, -2, -2)
		else
			paypal_amt = 0
		end if
	else
		paypal_amt = 0
	end if

%>
<!--#include virtual="/Connections/authnet.asp" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<script type="text/javascript">
	document.getElementById('addon-alert').style.display = 'none';
</script>
<!--#include file="functions/encrypt.asp"-->
<!--#include virtual="cart/generate_guest_id.asp"-->
<!--#include virtual="cart/inc_cart_main.asp"-->
<!--#include virtual="checkout/inc_order_details_google_analytics.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<!--#include virtual="cart/inc_cart_grandtotal.asp"-->



<%
'============================================================
if request.querystring("step") = "1" then

	if (session("state")  & "" & request.form("shipping-province-canada") & "" & session("shipping_province")) = "" then
		var_state = "-"
	else
		var_state = session("state")  & "" & request.form("shipping-province-canada") & "" & session("shipping_province")
	end if

	if session("country") = "Great Britain and Northern Ireland" then
		var_country = "United Kingdom"
	else
		var_country = session("country")
	end if
	
	' AUTHORIZATION AND CAPTURE REQUEST
	strAuthPayPal = "<?xml version=""1.0"" encoding=""utf-8""?>" _
	& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
	& MerchantAuthentication() _
	& "  <refId>" & session("invoiceid") & "</refId>" _
	& "  <transactionRequest>" _
	& "		<transactionType>authCaptureTransaction</transactionType>" _
	& "		<amount>" & paypal_amt & "</amount>" _
	& "  	<payment>" _
	& "  		<payPal>" _
	& "  			<successUrl>https://" & Request.ServerVariables("server_name") & "/checkout-paypal-authnet-v1.asp?step=2" _
	& "  			</successUrl>" _
	& "  			<cancelUrl>https://" & Request.ServerVariables("server_name") & "/cart.asp" _
	& "  			</cancelUrl>" _
	& "  			<paypalLc>US</paypalLc>" _	
	& "  			<paypalHdrImg>https://www.bodyartforms.com/images/bodyartforms-logo-text-black.png</paypalHdrImg>" _
	& "  			<paypalPayflowcolor>FFFFFF</paypalPayflowcolor>" _
	& "  		</payPal>" _
	& "  	</payment>" _
	& "  	<order><invoiceNumber>" & session("invoiceid") & "</invoiceNumber></order>" _	
	& "  	<shipTo>" _
	& "			<firstName>" & session("shipping_first") & "</firstName>" _
	& "			<lastName>" & session("shipping_last") & "</lastName>" _
	& "  	</shipTo>" _
	& "</transactionRequest>" _
	& "</createTransactionRequest>"
	Set objResponseAuthPayPal = SendApiRequest(strAuthPayPal)

	
	' PayPal request made
	If IsApiResponseSuccess(objResponseAuthPayPal) Then
	
		
		' if approved ...
		if objResponseAuthPayPal.selectSingleNode("/*/api:transactionResponse/api:responseCode").Text = "1" or objResponseAuthPayPal.selectSingleNode("/*/api:transactionResponse/api:responseCode").Text = "5" then
		
			var_message = "Re-directing to PayPal, please wait ..."
	
			session("paypal_transid") = objResponseAuthPayPal.selectSingleNode("/*/api:transactionResponse/api:transId").Text
			
			if request.cookies("OrderAddonsActive") = "" then
				' Update order with status and insert the PayPal transaction ID #
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE sent_items SET transactionID = ? WHERE ID = ?"
				objCmd.Parameters.Append(objCmd.CreateParameter("TransactionID",200,1,30,session("paypal_transid")))
				objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15,session("invoiceid")))
				objCmd.Execute()
			end if 	' if not OrderAddonsActive
		
			response.redirect objResponseAuthPayPal.selectSingleNode("/*/api:transactionResponse/api:secureAcceptance/api:SecureAcceptanceUrl").Text
		
		else
			' show message if not approved for 1st step to continue
			var_message = objResponseAuthPayPal.selectSingleNode("/*/api:transactionResponse").Text

			' WRITE TO PAYPAL LOG
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO tbl_checkout_logs (invoice_id, total, payflow_step, api_message, email, payment_method) VALUES (?,?,?,?,?,?)"
			objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, session("invoiceid")))
			objCmd.Parameters.Append(objCmd.CreateParameter("total",6,1,10, paypal_amt))
			objCmd.Parameters.Append(objCmd.CreateParameter("payflow_step",200,1,200, "step 1 - NOT APPROVED"))
			objCmd.Parameters.Append(objCmd.CreateParameter("api_message",200,1,2000, var_message))
			objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,100, session("email")))
			objCmd.Parameters.Append(objCmd.CreateParameter("payment_method",200,1,50, "PayPal"))
			objCmd.Execute()
		
		end if
		
	
	else	' AUTHORIZE ONLY
					
		var_message = objResponseAuthPayPal.selectSingleNode("/*/api:transactionResponse/api:errors/api:error/api:errorText").Text

		' WRITE TO PAYPAL LOG
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO tbl_checkout_logs (invoice_id, total, payflow_step, api_message, email, payment_method) VALUES (?,?,?,?,?,?)"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, session("invoiceid")))
		objCmd.Parameters.Append(objCmd.CreateParameter("total",6,1,10, paypal_amt))
		objCmd.Parameters.Append(objCmd.CreateParameter("payflow_step",200,1,200, "step 1 - ERROR"))
		objCmd.Parameters.Append(objCmd.CreateParameter("api_message",200,1,2000, var_message))
		objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,100, session("email")))
		objCmd.Parameters.Append(objCmd.CreateParameter("payment_method",200,1,50, "PayPal"))
		objCmd.Execute()
		
	end if



	
end if 'if request.querystring("step") = 1

'	----------------------------------------------------------------
if request.querystring("step") = 2 then 

		' AUTHORIZATION AND CAPTURE CONTINUE
		strCapturePayment = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<createTransactionRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <transactionRequest>" _
		& "		<transactionType>authCaptureContinueTransaction</transactionType>" _
		& "  	<payment>" _
		& "  		<payPal>" _
		& "  			<payerID>" & request.querystring("PayerID") & "</payerID>" _	
		& "  		</payPal>" _
		& "  	</payment>" _
		& "  	<refTransId>" & session("paypal_transid") & "</refTransId>" _
		& "</transactionRequest>" _
		& "</createTransactionRequest>"
		Set objResponseCapture = SendApiRequest(strCapturePayment)	
		
		If IsApiResponseSuccess(objResponseCapture) Then
					
			var_response_code = objResponseCapture.selectSingleNode("/*/api:transactionResponse/api:responseCode").Text
			
			var_raw_response_code = objResponseCapture.selectSingleNode("/*/api:transactionResponse/api:rawResponseCode").Text
		
		else	' PRIOR AUTHORIZATION CAPTURE
		
			var_message = objResponseCapture.selectSingleNode("/*/api:transactionResponse/api:errors/api:error/api:errorText").Text & " || " & session("paypal_transid") & " " & session("invoiceid")

			if request.querystring("PayerID") = "" then
				var_payerid = "no payer id found"
			else
				var_payerid = request.querystring("PayerID")
			end if

			' WRITE TO PAYPAL LOG
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO tbl_checkout_logs (invoice_id, payflow_step, api_message, payerID, transaction_id,total, email, payment_method) VALUES (?,?,?,?,?,?,?,?)"
			objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, session("invoiceid")))
			objCmd.Parameters.Append(objCmd.CreateParameter("payflow_step",200,1,200, "step 2 - ERROR"))
			objCmd.Parameters.Append(objCmd.CreateParameter("api_message",200,1,2000, var_message))
			objCmd.Parameters.Append(objCmd.CreateParameter("payerID",200,1,200, var_payerid))
			objCmd.Parameters.Append(objCmd.CreateParameter("transaction_id",200,1,200, session("paypal_transid")))
			objCmd.Parameters.Append(objCmd.CreateParameter("total",6,1,10, paypal_amt))
			objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,100, session("email")))
			objCmd.Parameters.Append(objCmd.CreateParameter("payment_method",200,1,50, "PayPal"))
			objCmd.Execute()
		
		end if




End If ' end step 2 ----------------------------------------

%>

<% If request.querystring("step") = 2 and var_response_code = 1 Then %>

	<div class="alert alert-success"> 
		<h5>Order Confirmation - Invoice # <%= Session("InvoiceID") %></h5>
		
		<strong>Your order has been approved. Thank you so much for shopping with us!</strong>
		<br/><br/>
		If you have any questions or concerns about your order or any other matter, please feel free to contact us <a class="alert-link" href="contact.asp">via e-mail</a> or by phone at (877) 223-5005.
		<br/><br/>
		<div class="mb-4">
		<a href="/index.asp" class="btn btn-outline-secondary">Back to Bodyartforms home page</a>
		</div>
		<a class="btn btn-lg btn-light border border-primary" href="https://g.page/r/CVDk_0MEUfIlEAQ/review" target="_blank" >
			<img src="/images/homepage/google-icon.png" class="mr-2" style="height: 50px" /> Review us on Google
		</a>
	</div>

	<% if request.querystring("feedback-submitted") = ""  then %>
	<!--
	<div class="card">
			<div class="card-header">
				<h6>Do you have a second to give us feedback about our website?</h6>
			</div>
				<div class="card-body">
					<form id="form-feedback" action="checkout_final.asp?feedback-submitted=yes" method="post">
						<div class="d-block mb-1">What device are you on?</div>
						
						<div class="custom-control custom-radio custom-control-inline">
							<input value="iPhone" type="radio" id="iPhone" name="platform" class="custom-control-input">
							<label class="custom-control-label" for="iPhone">iPhone</label>
						</div>
						<div class="custom-control custom-radio custom-control-inline">
							<input value="iPhone" type="radio" id="Android" name="platform" class="custom-control-input">
							<label class="custom-control-label" for="Android">Android</label>
						</div>
						<div class="custom-control custom-radio custom-control-inline">
							<input value="iPhone" type="radio" id="Tablet" name="platform" class="custom-control-input">
							<label class="custom-control-label" for="Tablet">Tablet</label>
						</div>
						<div class="custom-control custom-radio custom-control-inline">
							<input value="iPhone" type="radio" id="PC" name="platform" class="custom-control-input">
							<label class="custom-control-label" for="PC">PC</label>
						</div>
						<div class="custom-control custom-radio custom-control-inline">
							<input value="iPhone" type="radio" id="Mac" name="platform" class="custom-control-input">
							<label class="custom-control-label" for="Mac">Mac</label>
						</div>
						<div class="custom-control custom-radio custom-control-inline">
							<input value="iPhone" type="radio" id="Chromebook" name="platform" class="custom-control-input">
							<label class="custom-control-label" for="Chromebook">Chromebook</label>
						</div>
						<div class="form-group mt-3">
							<label for="feedback">Any comments or feedback?</label>
							<textarea class="form-control" name="feedback" id="feedback" rows="4"></textarea>
						</div>
						<button class="btn btn-purple" type="submit">Send feedback</button>
						<input type="hidden" name="email" value="<%= session("email") %>">
						<input type="hidden" name="name" value="<%= session("shipping_first") %>">
					</form>	
				</div>
		</div>
	-->
<% end if %>

<!--#include virtual="checkout/inc_random_code_generator.asp"--> 
<%

'ORDER IS ALREADY STORED ON /checkout/ajax_process_payment.asp and if Json returns PayPal then it goes to this page
%>
<!--#include virtual="checkout/inc_credits.asp" -->
<!--#include virtual="checkout/inc_use_discounts.asp"-->
<!--#include virtual="checkout/inc_giftcert_check_dupes.asp"--> 
<%
mailer_type = ""
%>
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<!--#include virtual="checkout/inc_giftcert_create.asp"--> 
<!--#include virtual="checkout/inc_wishlist_update.asp"--> 
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<!--#include virtual="cart/inc_cart_grandtotal.asp"-->
<%
if request.cookies("OrderAddonsActive") <> "" then
rs_getCart.ReQuery()
mailer_type = "addons approved"
%>
<!--#include virtual="checkout/inc-save-addon-items.asp" -->
<%
else 
mailer_type = "cc approved"
%>
<!--#include virtual="checkout/inc_deduct_quantities.asp" -->
<%
end if

payment_approved = "yes"
done_mailing_certs = "yes"
strCardType = "PayPal"
session("cc_status") = "approved" 

if request.cookies("OrderAddonsActive") = "" then

	' Update order with status and insert the PayPal transaction ID #
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET shipped = 'Review', ship_code = 'paid', transactionID = ? WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("TransactionID",200,1,30,session("paypal_transid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15,session("invoiceid")))
	objCmd.Execute()

	
	' WRITE TO CHECKOUT LOG
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_checkout_logs (invoice_id, payflow_step, api_message, total, email, payment_method) VALUES (?,?,?,?,?,?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, session("invoiceid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("payflow_step",200,1,200, "PAYMENT COMPLETE"))
	objCmd.Parameters.Append(objCmd.CreateParameter("api_message",200,1,2000, var_message))
	objCmd.Parameters.Append(objCmd.CreateParameter("total",6,1,10, paypal_amt))
		objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,100, session("email")))
		objCmd.Parameters.Append(objCmd.CreateParameter("payment_method",200,1,50, "PayPal"))
	objCmd.Execute()

%>
<!--#include virtual="checkout/inc-set-to-pending.asp" -->	
<% end if 'if not OrderAddonsActive
%>
<!--#include virtual="checkout/inc_items_to_email_array.asp"-->
<!--#include virtual="emails/function-send-email.asp"-->
<!--#include virtual="emails/email_variables.asp"-->
<!--#include virtual="checkout/inc_remove_items_from_cart.asp" -->
<!--#include virtual="checkout/inc_google_scripts.asp"--> 
<!--#include virtual="checkout/inc_remove_all_sessions_cookies.asp"-->
	<script type="text/javascript">
		// set cart count to 0
		$('#cart_count_text').html("");
		$('#mobile-cart-count').hide();
	</script>
<%
else ' Show decline information
%>
	<div class="alert alert-danger">
		<h5>PayPal Error</h5>
	<div class="my-2">
		Something went wrong with your payment. To try again <a class="alert-link" href="checkout.asp?type=paypal">click here</a> to go back to checkout.
	</div>
		<%= var_message %><br/>
		<%= var_raw_response_code %><br/>
		Authorize.net transaction ID# <%= session("paypal_transid") %>
	</div>
<%
end if ' If payment is a success
%>

<div style="height: 200px"></div>
<%
' WRITE TO CHECKOUT LOG
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "INSERT INTO tbl_checkout_logs (invoice_id, payflow_step, api_message, payerID, transaction_id,total, email, payment_method) VALUES (?,?,?,?,?,?,?,?)"
objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, session("invoiceid")))
objCmd.Parameters.Append(objCmd.CreateParameter("payflow_step",200,1,200, "BOTTOM OF PAGE"))
objCmd.Parameters.Append(objCmd.CreateParameter("api_message",200,1,2000, "DEBUGGER - CODE HAS PROCESSED THROUGH THE END OF THE PAGE"))
objCmd.Parameters.Append(objCmd.CreateParameter("payerID",200,1,200, 0))
objCmd.Parameters.Append(objCmd.CreateParameter("transaction_id",200,1,200, 0))
objCmd.Parameters.Append(objCmd.CreateParameter("total",6,1,10, 0))
objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,100, session("email")))
objCmd.Parameters.Append(objCmd.CreateParameter("payment_method",200,1,50, "PayPal"))
objCmd.Execute()
%>
<!--#include virtual="/bootstrap-template/footer.asp" -->