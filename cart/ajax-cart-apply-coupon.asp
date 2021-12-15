<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
'RESET applied coupons for debugging
'Session("CouponCode") = ""
'Session("GiftCertAmount") = 0
'Session("GiftCertCode") = ""

if session("textCouponBox") = "" then
	session("textCouponBox") = "Coupon or certificate "
end if

' Check to see if GIFT CERTIFICATE code is usable -------------------
If (Request("coupon_code") <> "" and Session("GiftCertAmount") = 0) or Session("GiftCertCode") <> "" Then
	
		if Request("coupon_code") <> "" then	
			var_code = Request("coupon_code")
		end if
		if Session("GiftCertCode") <> "" then
			var_code = Session("GiftCertCode")
		end if
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT ID, amount, code, invoice FROM dbo.TBLcredits WHERE code = ? AND amount > 0"
		objCmd.Parameters.Append(objCmd.CreateParameter("CertCode",200,1,40,var_code))
		Set rsGetGiftCert = objCmd.Execute()

		' STORE GIFT CERTIFICATE INTO SESSION VARIABLE
		If Not rsGetGiftCert.EOF Or Not rsGetGiftCert.BOF Then 
			Session("GiftCertAmount") = rsGetGiftCert.Fields.Item("amount").Value
			Session("GiftCertID") = rsGetGiftCert.Fields.Item("ID").Value
			Session("GiftCertCode") = rsGetGiftCert.Fields.Item("code").Value
			Session("GiftCertInvoice") = rsGetGiftCert.Fields.Item("invoice").Value
			cert_used = "yes"
		
			session("textCouponBox") = "Coupon" ' Change text box to enter code opposite to what's still available

		End If ' end Not rsCustomerCredit.EOF Or NOT rsCustomerCredit.BOF

		If rsGetGiftCert.EOF Or rsGetGiftCert.BOF Then ' if no gift certificate is found

			Session("GiftCertAmount") = 0
			Session("GiftCertID") = 0
			cert_used = "no"
			
			if Session("CouponCode") <> "" then
				session("textCouponBox") = "Certificate"
			end if				
			
		End if

		rsGetGiftCert.Close()
		Set rsGetGiftCert = Nothing

End If  'Check to see if GIFT CERTIFICATE code is usable -------------------



' CALCULATE COUPON DISCOUNT (IF ANY)-----------------------------------------------
If Request("coupon_code") <> "" and Session("CouponCode") = "" Then
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT DiscountID, DiscountDescription, DiscountCode, DiscountType, DiscountPercent, DateExpired, DateActive, BrandName, Clearance, ExcludeSaleItems FROM TBLDiscounts WHERE '" & now() & "' <= DateExpired AND '" & now() & "' >= DateActive AND ((DiscountCode = ? AND coupon_single_use = 0) OR (DiscountCode = ? AND coupon_single_use = 1 AND coupon__single_redeemed = 0)) ORDER BY DiscountPercent DESC"
		objCmd.Parameters.Append(objCmd.CreateParameter("Coupon1",200,1,40,Request("coupon_code")))
		objCmd.Parameters.Append(objCmd.CreateParameter("Coupon1",200,1,40,Request("coupon_code")))
		Set rsGetCoupon = objCmd.Execute()

		If Not rsGetCoupon.EOF Or Not rsGetCoupon.BOF Then ' If any coupons match

			Session("CouponCode") = rsGetCoupon.Fields.Item("DiscountCode").Value
			Session("CouponPercentage") = rsGetCoupon.Fields.Item("DiscountPercent").Value
			session("textCouponBox") = "Certificate" ' Change text box to enter code opposite to what's still available
			coupon_used = "yes"
			
			if rsGetCoupon.Fields.Item("BrandName").Value <> "None" then
				session("brand_coupon") = rsGetCoupon.Fields.Item("BrandName").Value
			else
				session("brand_coupon") = ""
			end if
		
		else
		
			Session("CouponCode") = ""
			Session("CouponPercentage") = ""
			session("brand_coupon") = ""
			coupon_used = "no"
			
			if Session("GiftCertAmount") <> 0 then
				session("textCouponBox") = "Coupon"
			end if
		
		End if				

end if ' Calculate COUPON CODE ------------------------------

' COUPON AND GIFT CERTIFICATE variables for displaying information on front end --------------

'if neither coupon NOR gift cert match then display to customer no discounts applied
if coupon_used = "no" and cert_used = "no" then
	discounts_applied = "no"
	
	if Len(Request("coupon_code")) <=20 then
		Valid_type = "COUPON"
	else
		Valid_type = "GIFT CERTIFICATE"
	end if
end if

'if the COUPON has already been matched, but the gift certificate does not then display no gift cert found
if Session("CouponCode") <> "" and cert_used = "no" then
	discounts_applied = "no"
	Valid_type = "GIFT CERTIFICATE"
end if

'if the GIFT CERT has already been matched, but the coupon does not then display no coupon found
if Session("GiftCertAmount") <> 0 and coupon_used = "no" then
	discounts_applied = "no"
	Valid_type = "COUPON"
end if

'if either coupon or gift cert match then display 
if coupon_used = "yes" or cert_used = "yes" then
	discounts_applied = "yes"
	
	if cert_used = "yes" then
		Valid_type = "GIFT CERTIFICATE"
	end if

	if coupon_used = "yes" then
		Valid_type = "COUPON"
	end if
end if

' END ----- COUPON AND GIFT CERTIFICATE variables for displaying information on front end --------------

if discounts_applied = "yes" and Request("coupon_code") <> "" then
%> 
	<div class="alert alert-success p-1"><%= Valid_type %> APPLIED</div>
<% end if %>

<%
if discounts_applied = "no" and Request("coupon_code") <> ""  then %> 
	<div class="alert alert-danger p-1"><%= Valid_type %> NOT VALID</div>
<% end if %>

<!--#include virtual="cart/inc_cart_main.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<!--#include virtual="cart/inc_cart_grandtotal.asp"-->	
<% if Session("CouponCode") <> "" then %>
	<div class="row">
		<div class="col-7">Coupon</div><div class="col-5">- $<span class="cart_coupon-amt"><%= FormatNumber(var_couponTotal, -1, -2, -2, -2) %></span></div>
	</div>
<% end if %>

<% if Session("GiftCertAmount") <> 0 then %>
	<div id="row_gift_cert">
		<div class="row">
			<div class="col-7">Gift certificate</div><div class="col-5">- <span id="cart_gift-cert"><%= FormatCurrency(var_total_giftcert_used, -1, -2, -2, -2) %></span></div>
		</div>
	</div>
<% ' if there is a gift certificate found
end if 
%>
<%

DataConn.Close()
Set DataConn = Nothing
%>