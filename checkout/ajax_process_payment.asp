<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% 
var_process_order = "yes"
check_stock = "yes"

if request.cookies("OrderAddonsActive") <> "" then
	var_addons_active = "yes"
end if


' BEGIN build { for .json return throughout this page}
%>
	{
<!--#include virtual="/template/inc_includes_ajax.asp" -->

<!--#include virtual="Connections/authnet.asp"-->
<!--#include virtual="functions/encrypt.asp"-->
<!--#include virtual="/cart/inc_cart_main.asp" -->
<% ' manually including stock check file below since variable var_process_order on inc_cart_main keeps a large chunk of that file out 
%>
<!--#include virtual="cart/inc_cart_stock_check.asp"-->
<%
' If no stock changes have occurred
if stock_display = "" then 
	var_stock_fail_json = "success"
	%>
		"stock_status":"success",
	<%
	'Let Google and Apple APIs know if order is flagged
	
	If session("flag") = "yes" Then%>
		"flagged":"yes",
	<%End If

else
	var_stock_fail_json = "fail"
	%>
		"stock_status":"fail"
	<%
end if 



' If no stock changes have occurred
if stock_display = "" then 

if cart_status = "not-empty" Then
%>

<!--#include virtual="checkout/inc_random_code_generator.asp"--> 
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<% rs_getCart.ReQuery() 
%>
<!--#include virtual="checkout/inc_store_shipping_selection.asp" -->
<!--#include virtual="cart/inc_cart_grandtotal.asp"-->
<%' ************************  DO NOT MOVE THE ORDER OF THESE LINKS AROUND BELOW. IT SCREWS STUFF UP **********************
%>
<!--#include virtual="checkout/inc_save_order.asp" -->
<!--#include virtual="checkout/inc_save_freeitems_to_order.asp"--> 
<% 
Set rsGetOrings = Nothing
if var_addons_active <> "yes" then %>
<!--#include virtual="functions/hash_extra_key.asp"-->
<!--#include virtual="accounts/inc_create_cim_account.asp"-->
<!--#include virtual="/functions/token.asp"-->
<!--#include virtual="accounts/inc_create_account.asp"-->
<!--#include virtual="checkout/inc_save_cim_profiles.asp" -->
<!--#include virtual="accounts/inc_update_cim_profiles.asp" -->
<!--#include virtual="checkout/inc_save_cims_to_order.asp" -->
<% end if %>

<% ' ************************************************

' Loop again to only store items after a successful payment 


' Set session total SPECIFICALLY FOR PAYPAL since it can't easily grab the shipping selection once it transfer to the Paypal checkout page}
	session("third_party_total") = var_grandtotal
%>
<!--#include virtual="checkout/inc_process_cashorder.asp"--> 
<!--#include virtual="checkout/inc_process_creditcard.asp"--> 
<!--#include virtual="checkout/inc_process_googlepay.asp"--> 
<!--#include virtual="checkout/inc_process_applepay.asp"--> 
<%

if payment_approved = "yes" then

	if var_addons_active <> "yes" then
		' Update order to set it "ON REVIEW"
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE sent_items SET shipped = 'Review', ship_code = 'paid' WHERE ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15,session("invoiceid")))
		objCmd.Execute()
	end if

	' WRITE TO CREDIT CARD LOG
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_checkout_logs (invoice_id, payflow_step, api_message, transaction_id,total, email, payment_method) VALUES (?,?,?,?,?,?,?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, session("invoiceid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("payflow_step",200,1,200, "PAYMENT COMPLETE"))
	objCmd.Parameters.Append(objCmd.CreateParameter("api_message",200,1,2000, "DEBUGGER - PAYMENT COMPLETE"))
	objCmd.Parameters.Append(objCmd.CreateParameter("transaction_id",200,1,200, 0))
	objCmd.Parameters.Append(objCmd.CreateParameter("total",6,1,10, var_grandtotal))
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,100, session("email")))
	objCmd.Parameters.Append(objCmd.CreateParameter("payment_method",200,1,50, "Credit card"))
	objCmd.Execute()
%>
<% if var_addons_active <> "yes" then %>
<!--#include virtual="/checkout/inc-set-to-pending.asp" -->
<!--#include virtual="checkout/inc_deduct_quantities.asp" -->
<% end if %>
<!--#include virtual="checkout/inc_credits.asp" -->
<!--#include virtual="checkout/inc_use_discounts.asp"-->
<%
if request.cookies("OrderAddonsActive") <> "" then
	mailer_type = "addons approved"
%>
<!--#include virtual="checkout/inc-save-addon-items.asp" -->
<%
end if
%>
<!--#include virtual="checkout/inc_giftcert_check_dupes.asp"--> 
<%
save_mailer_type = mailer_type
mailer_type = ""
%>
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<!--#include virtual="checkout/inc_giftcert_create.asp"--> 
<!--#include virtual="checkout/inc_wishlist_update.asp"--> 
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<!--#include virtual="checkout/inc_store_shipping_selection.asp" -->
<%
end if ' payment is approved

done_mailing_certs = "yes"
mailer_type = save_mailer_type
%>
<!--#include virtual="checkout/inc_items_to_email_array.asp"-->
<!--#include virtual="emails/function-send-email.asp"-->
<!--#include virtual="emails/email_variables.asp"-->
<%
if payment_approved = "yes" then
	' WRITE TO CREDIT CARD LOG
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_checkout_logs (invoice_id, payflow_step, api_message, transaction_id,total, email, payment_method) VALUES (?,?,?,?,?,?,?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, session("invoiceid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("payflow_step",200,1,200, "BOTTOM OF PAGE"))
	objCmd.Parameters.Append(objCmd.CreateParameter("api_message",200,1,2000, "DEBUGGER - CODE HAS PROCESSED THROUGH THE END OF THE PAGE"))
	objCmd.Parameters.Append(objCmd.CreateParameter("transaction_id",200,1,200, 0))
	objCmd.Parameters.Append(objCmd.CreateParameter("total",6,1,10, var_grandtotal))
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,100, session("email")))
	objCmd.Parameters.Append(objCmd.CreateParameter("payment_method",200,1,50, "Credit card"))
	objCmd.Execute()
end if


end if ' if Not rs_getCart.EOF
end if ' If no stock changes have occurred
%>
	}
<%
	' ^^ END build { for .json return throughout this page}
	
DataConn.Close()
Set DataConn = Nothing

%>