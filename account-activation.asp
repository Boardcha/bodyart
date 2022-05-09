<% @LANGUAGE="VBSCRIPT" %>
<%
page_title = "Account Activation"
page_description = "Bodyartforms Account Activation"
page_keywords = ""
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<%
email = Request("email")
activation_hash = Request("hash")

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM customers WHERE email = ? AND activation_hash = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,250,email))
objCmd.Parameters.Append(objCmd.CreateParameter("activation_hash",200,1,250,activation_hash))
Set rsGetUser = objCmd.Execute()
		
If Not rsGetUser.EOF Then
	Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
	password = "3uBRUbrat77V"
	data = rsGetUser.Fields.Item("customer_ID").Value
	encrypted = objCrypt.Encrypt(password, data)
	Response.Cookies("ID") = encrypted
	Response.Cookies("ID").Expires = DATE + 60
	Set objCrypt = Nothing	
End If	
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<div class="display-5 mb-3">
	Account Activation
</div>
<% If rsGetUser.EOF Then%>
	<div class="alert alert-danger">Invalid activation code. Please click on the link in the email sent upon registration.</div>
<%Else%>
	<%
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET active = 1 WHERE email = ? AND activation_hash = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,250,email))
	objCmd.Parameters.Append(objCmd.CreateParameter("activation_hash",200,1,250,activation_hash))
	objCmd.Execute()
	%>	
	<div class="alert alert-success">Your account is now active! Go to your <a href="account.asp">Account</a> page.</div>
<%End If%>
<%
If Not rsGetUser.EOF Then
	mailer_type = "new account"
	'Send mail out only when the user used the hash first time
	If rsGetUser("active") = false Then%>

	<!--#include virtual="/checkout/inc_random_code_generator.asp"-->
<!--#include virtual="/includes/inc-dupe-onetime-codes.asp"--> 
<%

' Prepare a one time use coupon for creating an account
var_cert_code = getPassword(15, extraChars, firstNumber, firstLower, firstUpper, firstOther, latterNumber, latterLower, latterUpper, latterOther)

' Call function
var_cert_code = CheckDupe(var_cert_code)

' Set extra mailer type
email_onetime_coupon = "yes"

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "INSERT INTO TBLDiscounts (DiscountCode, DateExpired, coupon_single_email, DiscountPercent, coupon_single_use, DateAdded, DiscountType, active, dateactive, coupon_assigned, DiscountDescription) VALUES (?, GETDATE()+30, ?, 10, 1, GETDATE(), 'Percentage', 'A', GETDATE()-1, 1, 'New account creation')"
objCmd.Parameters.Append(objCmd.CreateParameter("Code",200,1,30,var_cert_code))
objCmd.Parameters.Append(objCmd.CreateParameter("Email",200,1,30, email))
objCmd.Execute()

' Sent out account creation welcome email below

%>	
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/emails/email_variables.asp"-->	
	<%End If
End If%>	
<!--#include virtual="/bootstrap-template/footer.asp" -->