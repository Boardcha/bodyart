<%

' Process a cash order
' =================================================================================
if var_grandtotal > 0 and (request.form("cim_billing") = "cash" or request.form("cash") = "on") then

	'Set variable for mailer
	mailer_cash_order = "yes"
	session("cc_status") = "cash"

	' Send out cash order email below
	
%>
<!--#include virtual="checkout/inc_credits.asp" -->
<!--#include virtual="checkout/inc_use_discounts.asp"-->
<%
end if 

' Process an order that totals $0
' =================================================================================
if var_grandtotal <= 0 then

	payment_approved = "yes" ' will process on ajax_process_payment.asp
	session("cc_status") = "approved"
	
	'Set variable for mailer
	mailer_type = "cc approved"
	

	' Send out cash order email below
%>	
	"cc_approved":"no"
<% end if %>