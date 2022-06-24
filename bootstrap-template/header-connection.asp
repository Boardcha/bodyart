<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/functions/sha256.asp"-->
<!--#include virtual="/functions/salt.asp"-->
<%

referrer = Request.ServerVariables ("HTTP_REFERER")
If referrer <> "" _ 
	And Instr(referrer, "bodyartforms") = 0 _ 
	And Instr(referrer, "localhost") = 0 _ 
	And Instr(referrer, "127.0.0.1") = 0 _
	And Instr(referrer, "70.114.165.125") = 0 _ 
	And Instr(referrer, "75.109.218.58") = 0 _ 
	And Instr(referrer, "75.109.218.250") = 0 Then
		Session("referrer") = Left(referrer, 200)
End If	

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM customers WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10, CustID_Cookie))
Set rsGetUser = objCmd.Execute()
if request.cookies("ip-country") <> "" then
		strcountryName = request.cookies("ip-country")
else
		strcountryName = "US"
end if
' For my personal testing locally, always have it in sandbox mode so I don't fuck anythiing up
'session("sandbox") = "ON"

if request.cookies("adminuser") = "yes" then
		if request.querystring("inactive") = "yes" then
				session("inactive") = "yes"
		end if
		if request.querystring("inactive") = "no" then
				session("inactive") = ""
		end if
end if
        
         
' Check to see if there is a coupon that needs to be displayed
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "select website_text, DiscountDescription, DateActive, DateExpired, DiscountPercent, DiscountCode from TBLDiscounts where GETDATE() <= DateExpired AND GETDATE() >= DateActive AND show_on_website = 1"
Set rsDisplayCoupon = objCmd.Execute()
var_display_coupon_text = ""
var_display_coupon_code = ""
if NOT rsDisplayCoupon.EOF then
	var_display_coupon_amount = rsDisplayCoupon.Fields.Item("DiscountPercent").Value
	var_display_coupon_code = rsDisplayCoupon.Fields.Item("DiscountCode").Value 
	var_display_coupon_text = rsDisplayCoupon.Fields.Item("website_text").Value
	var_display_end_date = rsDisplayCoupon.Fields.Item("DateExpired").Value
end if

if session("exchange-rate") <> "" then
	exchange_rate = session("exchange-rate")
else
	exchange_rate = 1
end if

if session("exchange-symbol") <> "" then
	exchange_symbol = session("exchange-symbol")
else
	exchange_symbol = "$"
end if

' Reset if cookie = USD
if request.cookies("currency") = "USD" OR request.cookies("currency") = "" then
        exchange_rate = 1
        exchange_symbol = "$"
        session("exchange-rate") = ""
        session("exchange-symbol") = ""
        session("exchange-currency") = ""
end if

if request.querystring("secret") = "yes" then
        session("secret_sale") = "yes"
end if
        %>
<!DOCTYPE html>
<html lang="en">
<head>