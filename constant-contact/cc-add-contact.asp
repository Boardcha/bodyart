<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/functions/asp-json.asp"-->
<!--#include virtual="/functions/base64.asp" -->
<!--#include virtual="/Connections/constant-contact.asp" -->
<!--#include virtual="/constant-contact/cc-validate-token.asp"-->
<%
' =========== USE THE REFRESH TOKEN TO GET A NEW ACCESS TOKEN ======================
if cc_validate_error = "unauthorized" then
%>
<!--#include virtual="/constant-contact/cc-refresh-access-token.asp"-->
<%
end if

json_to_send = "{" & _
	"""email_address"":""" & request.querystring("email") & """," & _
	"""create_source"":""Footer signup""," & _
	"""list_memberships"": [" & _
		"""" & cc_baf_main_list_id & """" & _
	"]" & _
"}"

'response.write json_to_send


'======== ADDS OR UPDATES CONTACT BASED ON EMAIL ADDRESS ===========================
Set objAddContact = Server.CreateObject("MSXML2.ServerXMLHTTP")
objAddContact.open "POST", "https://api.cc.email/v3/contacts/sign_up_form", false
objAddContact.SetRequestHeader "Authorization", "Bearer " & cc_access_token
objAddContact.setRequestHeader "Content-Type", "application/json"
objAddContact.Send(json_to_send)

'response.write objAddContact.responseText

jsonContactString  = objAddContact.responseText
Set oJSON = New aspJSON
oJSON.loadJSON(jsonContactString)

cc_contact_status = oJSON.data("action") '==== RESPONSE WILL BE updated OR created

response.write replace(replace(jsonContactString,"[",""), "]", "")


'======Send out a one time coupon for a newly created subscriber (not an updated one) ========
if cc_contact_status = "created" then
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
	objCmd.Parameters.Append(objCmd.CreateParameter("Email",200,1,30, request.querystring("email") ))
	objCmd.Execute()
	
	' Sent out account creation welcome email below
%>	
	<!--#include virtual="/emails/function-send-email.asp"-->
	<!--#include virtual="/emails/email_variables.asp"-->
<%
end if '==== cc_contact_status = "created"


DataConn.Close()
Set DataConn = Nothing
%>