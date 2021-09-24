<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<% Response.ContentType = "application/json" 
%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="cart/generate_guest_id.asp"-->
<!--#include virtual="cart/inc_cart_main.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<!--#include virtual="checkout/inc_store_shipping_selection.asp"-->
<!--#include virtual="cart/ajax-sales-taxjar-rates.asp"-->
<!--#include virtual="cart/inc_cart_grandtotal.asp"-->
{  
	"subtotal":"<%= FormatNumber(var_subtotal, -1, -2, -2, -2) %>",
	"subtotal_after_discounts":"<%= FormatNumber(var_subtotal_after_discounts, -1, -2, -2, -2) %>",
	"couponamt":"<%= FormatNumber(var_couponTotal, -1, -2, -2, -2) %>",
	"preferred_discount":"<%= FormatCurrency(total_preferred_discount, -1, -2, -2, -2) %>",
	"salestax":"<%= FormatCurrency(session("amount_to_collect"), -1, -2, -2, -2) %>",
	"salestax_state":"<%= var_salestax_state %>",
	"salestax_forCalcs":"<%= FormatNumber(var_salesTax, -1, -2, -2, -2) %>",
	"salestax_rate":"<%= FormatNumber(var_salestax_rate, -1, -2, -2, -2) %>",
	"taxable_amount": "<%= session("taxable_amount") %>",
	"taxable_shipping": "<%= session("shipping_cost") %>",
	"amount_to_collect": "<%= session("amount_to_collect") %>",
	"state_tax_collectable": "<%= session("state_tax_collectable") %>",
	"county_tax_collectable": "<%= session("county_tax_collectable") %>",
	"city_tax_collectable": "<%= session("city_tax_collectable") %>",
	"special_district_tax_collectable": "<%= session("special_district_tax_collectable") %>",
	"fraudcheck_freegifts_subtotal":"<%= fraudcheck_freegifts_subtotal %>",
	"shipping":"<%= var_shipping_cost %>",
	"grandtotal":"<%= FormatNumber(var_grandtotal, -1, -2, -2, -2) %>",
	"shippingneeded":"<%= FormatCurrency(var_shipping_AmountNeeded, -1, -2, -2, -2) %>",
	"shippingfriendly":"<%= var_shipping_cost_friendly %>",
	"var_total_giftcert_used":"<%= FormatCurrency(var_total_giftcert_used, -1, -2, -2, -2) %>",
	"var_total_giftcert_dueback":"<%= FormatCurrency(var_total_giftcert_dueback, -1, -2, -2, -2) %>",
	"var_totalvalue_certs_incart":<%= var_totalvalue_certs_incart %>,
	"use_now_credits":"<%= FormatCurrency(credit_now, -1, -2, -2, -2) %>",
	"store_credit_amt":"<%= FormatCurrency(session("storeCredit_used"), -1, -2, -2, -2) %>",
	"store_credit_dueback":"<%= FormatCurrency(var_credit_due_todb, -1, -2, -2, -2) %>",
	"var_other_items":<%= var_other_items %>,
	"total_minus_tax":"<%= FormatNumber(var_grandtotal - var_salesTax, -1, -2, -2, -2) %>",
	"total_without_certsOrCredits":"<%= FormatNumber(var_total_without_certsOrCredits, -1, -2, -2, -2) %>",
	"total_without_shipping":"<%= FormatNumber(var_grandtotal - var_shipping_cost, -1, -2, -2, -2) %>",
	"weight":"<%= session("weight") %>"
}
<%
Set rs_getCart = Nothing
DataConn.Close()
Set DataConn = Nothing
%>