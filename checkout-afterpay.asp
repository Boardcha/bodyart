<%@LANGUAGE="VBSCRIPT"%>
<%
	page_title = "Bodyartforms AfterPay checkout"
	page_description = "Bodyartforms AfterPay checkout"
	page_keywords = "body jewelry, shopping cart, basket"
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<!--#include file="functions/encrypt.asp"-->
<!--#include virtual="cart/generate_guest_id.asp"-->
<!--#include virtual="cart/inc_cart_main.asp"-->
<!--#include virtual="checkout/inc_order_details_google_analytics.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<!--#include virtual="cart/inc_cart_grandtotal.asp"-->
<!--#include virtual="/functions/asp-json.asp"-->
<!--#include virtual="/Connections/afterpay-credentials.asp"-->


<%
'=============  This endpoint creates a checkout that is used to initiate the afterpay payment process. Afterpay uses the information in the checkout request to assist with the consumer’s pre-approval process. ========================================================================

Set objAfterPayCeckout = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
objAfterPayCeckout.open "POST", afterpay_url & "/payments/capture", false
objAfterPayCeckout.SetRequestHeader "Authorization", "Basic " & afterpay_api_credential & ""
objAfterPayCeckout.setRequestHeader "Accept", "application/json"
objAfterPayCeckout.setRequestHeader "Content-Type", "application/json"
objAfterPayCeckout.Send("{" & _
        """token"": """ & session("afterpay_checkout_token") & """," & _
        """merchantReference"": """ & session("invoiceid") & """" & _
    "}")

jsonCapturestring  = objAfterPayCeckout.responseText
Set oJSON = New aspJSON
oJSON.loadJSON(jsonCapturestring)

'response.write jsonCapturestring
'response.write "<br>TOKEN:" & session("afterpay_checkout_token")

If request.querystring("orderToken") = session("afterpay_checkout_token") Then

If oJSON.data("status") = "APPROVED" Then 
%>

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


<!--#include virtual="checkout/inc_random_code_generator.asp"--> 
<%
'Set array to store all order details (FOR CHECKOUT STORAGE INTO DATABASE)
	reDim array_details_2(8,0)
	Dim array_add_new : array_add_new = 0 
	
'ORDER IS ALREADY STORED ON /checkout/ajax_process_payment.asp and if Json returns PayPal then it goes to this page
%>
<!--#include virtual="checkout/inc_credits.asp" -->
<!--#include virtual="checkout/inc_use_discounts.asp"-->
<!--#include virtual="checkout/inc_giftcert_check_dupes.asp"--> 
<%
mailer_type = ""
rs_getCart.ReQuery() 
%>
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->

<!--#include virtual="checkout/inc_giftcert_create.asp"--> 
<!--#include virtual="checkout/inc_wishlist_update.asp"--> 
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<!--#include virtual="cart/inc_cart_grandtotal.asp"-->

<!--#include virtual="checkout/inc_deduct_quantities.asp" -->

<%

payment_approved = "yes"
done_mailing_certs = "yes"
mailer_type = "cc approved"
strCardType = "Afterpay"
session("cc_status") = "approved" 
pay_method_afterpay = "yes"


	' Update order with status and insert the PayPal transaction ID #
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE sent_items SET shipped = ?, ship_code = 'paid', transactionID = ? WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("status",200,1,30,"Review"))
	objCmd.Parameters.Append(objCmd.CreateParameter("TransactionID",200,1,30, oJSON.data("id") ))
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, session("invoiceid") ))
	objCmd.Execute()
%>
<!--#include virtual="emails/function-send-email.asp"-->
<!--#include virtual="emails/email_variables.asp"-->
<%	
rs_getCart.ReQuery() ' for Google
%>
<!--#include virtual="/checkout/inc-set-to-pending.asp" -->	
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
		<h5>Afterpay Card Declined</h5>
	<div class="my-2">
		Afterpay declined your credit card information. To try again <a class="alert-link" href="checkout.asp?type=afterpay">click here</a> to go back to checkout.
	</div>
	</div>
<%
end if ' If payment is a success

else ' if tokens do not match
%>
<div class="alert alert-danger">
	<h5>Afterpay Error</h5>
<div class="my-2">
	Checkout tokens did not match. To try again <a class="alert-link" href="checkout.asp?type=afterpay">click here</a> to go back to checkout.
</div>
</div>
<% end if '=== if tokens match %>



<div style="height: 200px"></div>
<!--#include virtual="/bootstrap-template/footer.asp" -->