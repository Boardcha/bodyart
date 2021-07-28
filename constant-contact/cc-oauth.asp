<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/functions/asp-json.asp"-->
<!--#include virtual="/functions/base64.asp" -->
<!--#include virtual="/Connections/constant-contact.asp" -->
<a href="https://api.cc.email/v3/idfed?client_id=1619363b-10b8-42ed-8e8a-9f8f35484c9a&redirect_uri=https%3A%2F%2Flocalhost%2Fconstant-contact%2Fcc-oauth.asp&response_type=code&scope=contact_data">Sign in to request access token</a>
<br/>
<br/>
<%
' -------- CONSTANT CONTACT VERSION 3 API CONNECTION ----------------
' -------- ONLY USE THIS PAGE TO GENERATE THE INITIAL ACCESS AND REFRESH TOKEN SO THAT WE CAN USE THE API UNDER OUR ACCOUNT ----------------
' ====== ACCESS TOKENS LAST 24 HOURS. WITH EACH ACCESS TOKEN WE ARE GIVEN A REFRESH TOKEN
' ====== USE THE REFRESH TOKEN TO GET NEW 24 ACCESS TOKENS. IF YOU DON'T USE THE LAST REFRESH TOKEN YOU HAVE TO USE THIS PAGE AGAIN TO RE-AUTHORIZE AND GET A NEW SET

' ======= INITIAL AUTHORIZATION - REQUIRES LOGGING INTO CONSTANT CONTACT AND ALLOWING ACCESS
' ======= REDIRECT URL MUST BE SET IN CONSTANT CONTACT WEBSITE FOR SECURITY. This is set in the My Applications section under the API Docs =================

'url = "https://api.cc.email/v3/idfed?client_id=" & cc_api_key_client_id & "&redirect_uri=https%3A%2F%2Flocalhost%2Fconstant-contact%2Fcc-oauth.asp&response_type=code&scope=contact_data"
'Set objGetAccessToken = Server.CreateObject("MSXML2.ServerXMLHTTP")
'objGetAccessToken.open "GET", url, false
'objGetAccessToken.Send
'response.write objGetAccessToken.responseText

if request.querystring("code") <> "" then

	' AUTHORIZATIOIN CODE RESPONSE TO GET ACCESS TOKEN (LASTS 24 HOURS)
	url = "https://idfed.constantcontact.com/as/token.oauth2?code=" & request.querystring("code") & "&redirect_uri=https%3A%2F%2Flocalhost%2Fconstant-contact%2Fcc-oauth.asp&grant_type=authorization_code&scope=contact_data"
	Set objGetAccessToken = Server.CreateObject("MSXML2.ServerXMLHTTP")
	objGetAccessToken.open "POST", url, false
	objGetAccessToken.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objGetAccessToken.SetRequestHeader "Authorization", "Basic " & cc_auth_encrypted
	objGetAccessToken.Send

	jsonTokenString  = objGetAccessToken.responseText
	Set oJSON = New aspJSON
	oJSON.loadJSON(jsonTokenString)

	'response.write jsonTokenString

	if oJSON.data("error") <> "invalid_grant" then

		cc_new_access_token = oJSON.data("access_token")
		cc_new_refresh_token = oJSON.data("refresh_token")

		'response.write "<br/>access token: " & cc_new_access_token
		'response.write "<br/>refresh token: " & cc_new_refresh_token

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "USE sandbox UPDATE tbl_constant_contact_api_keys SET access_token = ?, refresh_token = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("access_token",200,1,200, cc_new_access_token))
		objCmd.Parameters.Append(objCmd.CreateParameter("refresh_token",200,1,200, cc_new_refresh_token))
		Set rsKeys = objCmd.Execute()

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "USE baf_site UPDATE tbl_constant_contact_api_keys SET access_token = ?, refresh_token = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("access_token",200,1,200, cc_new_access_token))
		objCmd.Parameters.Append(objCmd.CreateParameter("refresh_token",200,1,200, cc_new_refresh_token))
		Set rsKeys = objCmd.Execute()

	end if '----- if token is created

end if


DataConn.Close()
Set DataConn = Nothing
%>
