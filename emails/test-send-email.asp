<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/cart/inc_cart_main.asp" -->
	<%
	reDim array_details_2(12,0)
	Dim array_add_new : array_add_new = 0 
%>

	<!--#include virtual="/cart/inc_cart_loopitems-begin.asp"-->
	<!--#include virtual="/checkout/inc_orderdetails_toarray.asp"--> 
<!--#include virtual="/cart/inc_cart_loopitems-end.asp"-->
<!--#include virtual="/checkout/inc_freeitems_toarray.asp"--> 

<%
done_mailing_certs = "no"
'payment_approved = "yes"
'mailer_cash_order = "yes"

'rec_email = "amanda.bodyartforms@gmail.com"
%>
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/emails/email_variables.asp"-->
<%
DataConn.Close()
Set DataConn = Nothing
%>