<%
' This page looks for a valid token in the DB. Tokens are valid only 1 hour. If it is valid, it uses that token. If it is expired, refreshes the token from Etsy.
' The token will be assigned into the variable etsy_access_token varibale. This page should be included at the top of related pages. So, you will have a valid "etsy_access_token".
%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/chilkat.asp" -->
<!--#include virtual="/Connections/etsy-constants.asp" -->
<%
SqlString = "SELECT * FROM TBL_Access_Tokens WHERE provider = 'etsy-access-token' AND GETDATE() < date_expires" 
Set rsToken = DataConn.Execute(SqlString)
If Not rsToken.EOF Then
	'If it is valid
	etsy_access_token = rsToken("access_token")
Else 'If it is expired get a new token
	SqlString = "SELECT * FROM TBL_Access_Tokens WHERE provider = 'etsy-refresh-token'" 
	Set rsRefreshToken = DataConn.Execute(SqlString)	
	If Not rsRefreshToken.EOF then
		set json = Server.CreateObject("Chilkat_9_5_0.JsonObject")
		set req = Server.CreateObject("Chilkat_9_5_0.HttpRequest")
		req.AddParam "grant_type", "refresh_token"
		req.AddParam "client_id", etsy_consumer_key
		req.AddParam "refresh_token", rsRefreshToken("access_token")

		set http = Server.CreateObject("Chilkat_9_5_0.Http")

		' resp is a Chilkat_9_5_0.HttpResponse
		Set resp = http.PostUrlEncoded("https://api.etsy.com/v3/public/oauth/token", req)
		If (http.LastMethodSuccess <> 1) Then
			Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
			Response.End
		End If

		'  Load the JSON response.
		success = json.Load(resp.BodyStr)
		json.EmitCompact = 0

		'  If the response status code is not 200, then it's an error.
		If (resp.StatusCode <> 200) Then
			Response.End
		End If
		
		etsy_access_token = json.StringOf("access_token")
		etsy_refresh_token = json.StringOf("refresh_token")
		etsy_token_expiration = json.StringOf("expires_in")

		If etsy_access_token<>"" AND etsy_refresh_token<>"" AND etsy_token_expiration<>"" Then		
			SqlString = "DELETE FROM TBL_Access_Tokens WHERE provider = 'etsy-access-token' OR provider = 'etsy-refresh-token'" 
			DataConn.Execute(SqlString)		
			
			SqlString = "INSERT INTO TBL_Access_Tokens (access_token, provider, date_expires) VALUES('" & etsy_access_token & "', 'etsy-access-token', DATEADD(ss," & etsy_token_expiration & ", GETDATE()))" 
			DataConn.Execute(SqlString)	
			
			SqlString = "INSERT INTO TBL_Access_Tokens (access_token, provider) VALUES('" & etsy_refresh_token & "', 'etsy-refresh-token')" 
			DataConn.Execute(SqlString)	
		End If 
	Else
		Response.Write "Refresh token could not be found in the DB. To get one, call the page ""etsy-accees-token.asp"" and grant access."
		Response.End
	End If
End If

Set rsToken = Nothing
Set rsRefreshToken = Nothing
Set req = Nothing
Set json = Nothing
Set http = Nothing
Set resp = Nothing
%>