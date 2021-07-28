<%
' -------- CONSTANT CONTACT VERSION 3 API CONNECTION ----------------
' -------- THIS PAGE USES THE REFRESH TOKEN TO GET A NEW 24 ACCESS TOKEN ----------------


	' SEND REFRESH TOKEN TO GET NEW ACCESS TOKEN (LASTS 24 HOURS)
	url = "https://idfed.constantcontact.com/as/token.oauth2?refresh_token=" & cc_refresh_token & "&grant_type=refresh_token"
	Set objRefreshToken = Server.CreateObject("MSXML2.ServerXMLHTTP")
	objRefreshToken.open "POST", url, false
	objRefreshToken.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objRefreshToken.SetRequestHeader "Authorization", "Basic " & cc_auth_encrypted
	objRefreshToken.Send

	jsonRefreshString  = objRefreshToken.responseText
	Set oJSON = New aspJSON
	oJSON.loadJSON(jsonRefreshString)

	'response.write "Refresh token:<br>" & jsonRefreshString & "<br/><br>"


	if oJSON.data("error") <> "invalid_grant" then

		cc_access_token = oJSON.data("access_token")
		cc_refresh_token = oJSON.data("refresh_token")

		'response.write "<br/>access token: " & cc_access_token
		'response.write "<br/>refresh token: " & cc_refresh_token

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "USE sandbox UPDATE tbl_constant_contact_api_keys SET access_token = ?, refresh_token = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("access_token",200,1,200, cc_access_token))
		objCmd.Parameters.Append(objCmd.CreateParameter("refresh_token",200,1,200, cc_refresh_token))
		Set rsKeys = objCmd.Execute()

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "USE baf_site UPDATE tbl_constant_contact_api_keys SET access_token = ?, refresh_token = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("access_token",200,1,200, cc_access_token))
		objCmd.Parameters.Append(objCmd.CreateParameter("refresh_token",200,1,200, cc_refresh_token))
		Set rsKeys = objCmd.Execute()

	end if '----- if token is created


%>
