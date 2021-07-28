<%
' Get expirdation date of oAuth access token

Set objCheckToken = Server.CreateObject("MSXML2.ServerXMLHTTP")
objCheckToken.open "POST", "https://api.cc.email/v3/token_info", false
objCheckToken.setRequestHeader "Content-Type", "application/json"
objCheckToken.SetRequestHeader "Authorization", "Bearer " & cc_access_token
objCheckToken.Send("{" & _
	"""token"":" & cc_access_token & "" & _
	"}")

jsonCheckTokenString  = objCheckToken.responseText
Set oJSON = New aspJSON
oJSON.loadJSON(jsonCheckTokenString)

'response.write "Check token:<br>" & jsonCheckTokenString & "<br/><br>"

cc_validate_error = oJSON.data("error_key")
'response.write "error (if any): " & cc_validate_error & "<br/><br/>"
%>