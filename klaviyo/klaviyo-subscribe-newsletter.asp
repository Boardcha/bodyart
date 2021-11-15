<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/Connections/klaviyo.asp" -->
<!--#include virtual="/Connections/chilkat.asp" -->

<%
var_email = request("email")
'=====  UTEZqk is the list ID for the basic Newsletter

' =================== CHECK TO SEE IF EMAIL SIGNUP IS CURRENTLY ALREADY IN THE LIST AND IF NOT SIGN THEM UP 
'  Connect to the REST server.
set rest = Server.CreateObject("Chilkat_9_5_0.Rest")
bTls = 1
port = 443
bAutoReconnect = 1
success = rest.Connect("https://a.klaviyo.com/",port,bTls,bAutoReconnect)
success = rest.AddHeader("Content-Type","application/json")
success = rest.AddHeader("Accept","application/json")
If (success = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If

	options = "{" & _
	" ""api_key"" : """ & klaviyo_private_key & """," & _
    " ""emails"" : [ " & _
        " """ & var_email &""" " & _
    "]}"
	set http = Server.CreateObject("Chilkat_9_5_0.Http")
	http.SetRequestHeader "Content-Type", "application/json"
	http.Accept = "application/json"
	
   
	Set resp = http.PostJson2("https://a.klaviyo.com/api/v2/list/UTEZqk/get-members?","application/json", options)
	If (http.LastMethodSuccess = 0) Then
		Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
		Response.End
	End If
	jsonResponseStr = resp.BodyStr
	response.write jsonResponseStr
	'Response.Write "<br>Find member result: " &  jsonResponseStr

	'======= IF MEMBER IS NOT FOUND ... KLAVIYO RETURNS [] AS RESULT ... THEN SEND OUT THE ONE TIME COUPON ================
	if jsonResponseStr = "[]" then
%>
		<!--#include virtual="/checkout/inc_random_code_generator.asp"-->
		<!--#include virtual="/includes/inc-dupe-onetime-codes.asp"--> 
<%
		' Prepare a one time use coupon for creating an account
		var_cert_code = getPassword(15, extraChars, firstNumber, firstLower, firstUpper, firstOther, latterNumber, latterLower, latterUpper, latterOther)
		
		' Call function
		var_cert_code = CheckDupe(var_cert_code)

		' Set extra mailer type
		email_newsletter_signup_coupon = "yes"

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO TBLDiscounts (DiscountCode, DateExpired, coupon_single_email, DiscountPercent, coupon_single_use, DateAdded, DiscountType, active, dateactive, coupon_assigned, DiscountDescription) VALUES (?, GETDATE()+30, ?, 15, 1, GETDATE(), 'Percentage', 'A', GETDATE()-1, 1, 'Newsletter signup')"
		objCmd.Parameters.Append(objCmd.CreateParameter("Code",200,1,30,var_cert_code))
		objCmd.Parameters.Append(objCmd.CreateParameter("Email",200,1,30, var_email ))
		objCmd.Execute()

		' Sent out account creation welcome email below
%>
		<!--#include virtual="/emails/function-send-email.asp"-->
		<!--#include virtual="/emails/email_variables.asp"-->
<%
	' =================== SUBSCRIBE NEW MEMBER TO NEWSLETTER
	'  Connect to the REST server.
	set rest = Server.CreateObject("Chilkat_9_5_0.Rest")
	bTls = 1
	port = 443
	bAutoReconnect = 1
	success = rest.Connect("https://a.klaviyo.com/",port,bTls,bAutoReconnect)
	success = rest.AddHeader("Content-Type","application/json")
	success = rest.AddHeader("Accept","application/json")
	If (success = 0) Then
		Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
		Response.End
	End If

		options = "{" & _
		" ""profiles"" : [ " & _
			"{ ""email"" : """ & var_email &""" }" & _
		"]}"
		set http = Server.CreateObject("Chilkat_9_5_0.Http")
		http.SetRequestHeader "Content-Type", "application/json"
		http.Accept = "application/json"
		
	
		Set resp = http.PostJson2("https://a.klaviyo.com/api/v2/list/UTEZqk/subscribe?api_key=" & klaviyo_private_key ,"application/json", options)
		If (http.LastMethodSuccess = 0) Then
			Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
			Response.End
		End If
		jsonResponseStr = resp.BodyStr
		'Response.Write "<br>Subscribe member result: " & jsonResponseStr

	end if '===== If member is not found =======================================================

	DataConn.Close()
%>