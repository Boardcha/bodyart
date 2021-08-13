<%@LANGUAGE="VBSCRIPT"%>
<%
' This page obtains a new access token from Etsy, writes it to the TBL_Access_Tokens table with the expiration time of the token.
' If the acess token is expired, the token will be refreshed using the etsy-refresh-token.asp page automnatically when it is expired.
%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/chilkat.asp" -->
<!--#include virtual="/admin/etsy/etsy-constants.asp" -->
<%
If Request.ServerVariables("SERVER_NAME")="localhost" Then strProtocol = "http://" Else strProtocol = "https://"
redirection_url = strProtocol & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME") 
	
etsy_access_token = ""
stateReceived = Request.QueryString("state")

' Check to see if this is our redirection from Etsy containing the access token.
If stateReceived <> "" then
	' Make sure this is the redirect for our session.
	If stateReceived <> Session("oauth2_state") then
		etsy_access_token = "invalid_state"
	ElseIf Request.QueryString("code") <> "" then
		' Exchange authorization code for access tokens
		set http = Server.CreateObject("Chilkat_9_5_0.Http")	
		set req = Server.CreateObject("Chilkat_9_5_0.HttpRequest")
		req.AddParam "client_id", etsy_consumer_key
		req.AddParam "grant_type","authorization_code"
		req.AddParam "code_verifier", Session("code_verifier")
		req.AddParam "code",Request.QueryString("code")
		req.AddParam "redirect_uri", redirection_url
		
		' resp is a Chilkat_9_5_0.HttpResponse
		Set resp = http.PostUrlEncoded("https://api.etsy.com/v3/public/oauth/token",req)
		If (http.LastMethodSuccess <> 1) Then
			Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
			Response.End
		End If

		set json = Server.CreateObject("Chilkat_9_5_0.JsonObject")
		json.EmitCompact = 0
		json.Load(resp.BodyStr)
		
		etsy_access_token = json.StringOf("access_token")
		etsy_token_expiration = json.StringOf("expires_in")
		etsy_refresh_token = json.StringOf("refresh_token")
		
		'Overwrite access and refresh tokens
		SqlString = "DELETE FROM TBL_Access_Tokens WHERE provider = 'etsy-access-token' OR provider = 'etsy-refresh-token'" 
		DataConn.Execute(SqlString)	
		
		SqlString = "INSERT INTO TBL_Access_Tokens (access_token, provider, date_expires) VALUES('" & etsy_access_token & "', 'etsy-access-token', DATEADD(ss," & etsy_token_expiration & ", GETDATE()))" 
		DataConn.Execute(SqlString)	
		
		SqlString = "INSERT INTO TBL_Access_Tokens (access_token, provider, date_expires) VALUES('" & etsy_refresh_token & "', 'etsy-refresh-token', DATEADD(ss," & etsy_token_expiration & ", GETDATE()))" 
		DataConn.Execute(SqlString)			
		
		Response.Write "All good!" & "<br>"
		Response.Write "Access token and refresh token saved to the DB. Please do not use this page again to get new tokens, this page requires ""GRANT ACCESS"" each time. BAF Admin interface will use etsy-refresh-token.asp page to get new access tokens automatically, when they are expired." & "<br>"
		Response.Write "If you accidently delete refresh token from the DB, then you can call this page (etsy-access-token.asp) and grant access to store a new refresh token in the DB." & "<br>"
		Response.End    

	End if
Else
	' If stateReceived is empty, initiate requesting access token
	'code_verifier = generateVerifier()
	code_verifier = "mXJsm4Cjt4D5vgbk_AKpJUKHuyvcbWFvb8joE3woHfY"
	Session("code_verifier") = code_verifier
	'code_challenge = getCodeChallenge(code_verifier)
	code_challenge = "8LZnPaVIN_b2Cib9SK6ly5PL_ciEugR9v5qY8ooeE0k"
	Session("code_challenge") = code_challenge
	set req = Server.CreateObject("Chilkat_9_5_0.HttpRequest")
	call req.AddParam("client_id", etsy_consumer_key)
	' This will redirect to this same ASP page..
	call req.AddParam("redirect_uri", redirection_url)
	call req.AddParam("response_type", "code")
	'call req.AddParam("scope", "address_r address_w billing_r cart_r cart_w email_r favorites_r favorites_w	feedback_r listings_d listings_r listings_w profile_r profile_w recommend_r recommend_w shops_r shops_w transactions_r transactions_w")
	call req.AddParam("scope", "address_r address_w billing_r cart_r cart_w email_r favorites_r favorites_w feedback_r listings_d listings_r listings_w profile_r profile_w recommend_r recommend_w shops_r shops_w transactions_r transactions_w")
	call req.AddParam("code_challenge_method", "S256")
	call req.AddParam("code_challenge", code_challenge)
	' Random state string
	stateToGo = RandomNumber(8)
	call req.AddParam("state", stateToGo)
	Session("oauth2_state") = stateToGo
	
	auth_url = "https://www.etsy.com/oauth/connect?" + req.GetUrlEncodedParams()
	Response.Redirect auth_url
	
End if			

set rsToken = nothing
DataConn.Close()	
%>	
<%
Function SHA256(sText)
	set crypt = Server.CreateObject("Chilkat_9_5_0.Crypt2")
	crypt.HashAlgorithm = "sha256"
	crypt.Charset = "utf-8"
	hashBytes = crypt.HashString(sText)

	set sb = Server.CreateObject("Chilkat_9_5_0.StringBuilder")
	success = sb.AppendEncoded(hashBytes, "base64")
	hash = sb.GetAsString()
	SHA256 = hash
End Function

Function StringToByteArray(s)
  Dim i, byteArray
  For i=1 To Len(s)
    byteArray = byteArray & ChrB(Asc(Mid(s,i,1)))
  Next
  StringToByteArray = byteArray
End Function

Function generateVerifier()
	Session("rnd") = RandomNumber(32)
	generateVerifier = Base64Encode(Session("rnd"))
End Function

Function getCodeChallenge(code_verifier)
	getCodeChallenge = Base64Encode(SHA256(code_verifier))
End Function

Function RandomNumber(length)
	Randomize
	For i = 1 to length
		temp = temp & Int(Rnd * 9) + 1
	Next
	RandomNumber = temp
End Function

Function Base64Encode(sText)
	Dim oXML, oNode
	Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
	Set oNode = oXML.CreateElement("base64")
	oNode.dataType = "bin.base64"
	oNode.nodeTypedValue = Stream_StringToBinary(sText)
	Base64Encode = oNode.text

	Base64Encode = Replace(Base64Encode, "=", "")
	Base64Encode = Replace(Base64Encode, "+", "-")
	Base64Encode = Replace(Base64Encode, "/", "_")

	Set oNode = Nothing
	Set oXML = Nothing
End Function

Function Base64Encode2(sText)
	set bd = Server.CreateObject("Chilkat_9_5_0.BinData")
	success = bd.AppendString(sText,"utf-8")
	Base64Encode2 = bd.GetEncoded("base64")
End Function

Private Function Stream_StringToBinary(Text)
	Const adTypeText = 2
	Const adTypeBinary = 1
	Dim BinaryStream 'As New Stream
	Set BinaryStream = CreateObject("ADODB.Stream")
	BinaryStream.Type = adTypeText
	BinaryStream.CharSet = "us-ascii"
	BinaryStream.Open
	BinaryStream.WriteText Text
	BinaryStream.Position = 0
	BinaryStream.Type = adTypeBinary
	BinaryStream.Position = 0
	Stream_StringToBinary = BinaryStream.Read
	Set BinaryStream = Nothing
End Function

%>