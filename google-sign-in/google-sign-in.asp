<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/Connections/authnet.asp"-->
<!--#include virtual="/Connections/chilkat.asp" -->
<!--#include virtual="/Connections/google-oauth-credentials.inc" -->
<!--#include virtual="/functions/hash_extra_key.asp"-->
<!--#include virtual="/functions/encrypt.asp"-->
<!--#include virtual="/functions/token.asp"-->
<!--#include virtual="/functions/salt.asp"-->
<!--#include virtual="/functions/sha256.asp"-->
<%

g_csrf_token_cookie = Request.Cookies("g_csrf_token")
g_csrf_token_post = Request.Form("g_csrf_token")
idToken = Request.Form("credential")

set jwt = Server.CreateObject("Chilkat_9_5_0.Jwt")
payload = jwt.GetPayload(idToken)
set json = Server.CreateObject("Chilkat_9_5_0.JsonObject")
success = json.Load(payload)
json.EmitCompact = 0

google_aud = json.StringOf("aud")
google_iss = json.StringOf("iss")
google_email = json.StringOf("email")
google_firstName = json.StringOf("given_name")
google_lastName = json.StringOf("family_name")
google_user_id = json.StringOf("sub")


'Check if the info provided by JS and the one comes from Google endpoint is matched
'aud and iss check is necessary to prevent ID tokens issued to a malicious app being used to access data about the same user on the server.

If g_csrf_token_cookie = g_csrf_token_post AND Instr(google_aud, google_oauth_clientId) > 0 AND (google_iss = "accounts.google.com" OR google_iss = "https://accounts.google.com") AND google_user_id <> "" Then 'On successful sign-in

	google_signin_email = google_email
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT customer_ID, email FROM customers WHERE google_user_id = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("google_user_id",200,1,50, google_user_id))
	set rsGoogle = objCmd.Execute()
	If Not rsGoogle.EOF Then ' If google_user_id exists
		doLogin(rsGoogle("customer_ID"))
	ElseIf emailExist() Then 'If email exist and google_user_id doesn't, connect them.
		existentCustomerId = updateGoogleIdForExistentEmail() 
		doLogin(existentCustomerId)
	Else
		'Create a new user
		%>
		<!--#include virtual="accounts/inc_create_account.asp"-->
		<%
		doLogin(var_our_custid)
	End If
Else
	%>
	{ "status":"logged-out" }
	<%	
	
End If
' ==== END OF PAGE ====


' ==== LOCAL FUNCTIONS ====
Function emailExist()
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT customer_ID, email FROM customers WHERE email = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,50, google_signin_email))
	set rsEmailFound = objCmd.Execute()
	If rsEmailFound.EOF Then emailExist = false Else emailExist = true
	Set rsEmailFound = Nothing
End Function 

Function updateGoogleIdForExistentEmail()
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET google_user_id = ? WHERE email = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("google_user_id",200,1,200,google_user_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,200,google_signin_email))
	objCmd.Execute()
	
	'Retrieve customer ID number from our database
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT customer_ID, email FROM customers WHERE email = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,200,google_signin_email))
	Set rsGetUserID = objCmd.Execute()
	updateGoogleIdForExistentEmail = rsGetUserID("customer_ID")
End Function 

Sub doLogin(customerId)

	session("login_email") = google_signin_email

	' Write last login date
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET last_login = '" & now() & "' WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,customerId))
	objCmd.Execute()
	
	' Set session variable to modify shipping/billing information in account
	session("custID_account") = customerId
	var_our_custid = customerId
	Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")

	password = "3uBRUbrat77V"
	data = customerId

	encrypted = objCrypt.Encrypt(password, data)
	Response.Cookies("ID") = encrypted
	Response.Cookies("ID").Expires = DATE + 30

	Set objCrypt = Nothing

	if Request.Cookies("ID") <> "" then 
		%>
		{ "status":"logged-in" }	
		<!--#include virtual="/accounts/inc_transfer_cart_contents.asp" -->
		<%
		' decrypt customer ID cookie
		Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")

		password = "3uBRUbrat77V"
		data = request.Cookies("ID")

		If len(data) > 5 then ' if
			decrypted = objCrypt.Decrypt(password, data)
		end if

		  if data <> decrypted then
			  CustID_Cookie = decrypted
		  else
			  CustID_Cookie = 0
		  end if

		Set objCrypt = Nothing
		%>
		<!--#include virtual="/cart/inc_cart_main.asp"-->
		<%	
	end if  	
End Sub

Set rsGoogle = Nothing
%>
