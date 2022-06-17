<%
Function GetCurrentUrl()
	Dim strProtocol
	Dim strDomain
	Dim strPath
	Dim strQueryString
	Dim strFullUrl

	If lcase(Request.ServerVariables("HTTPS")) = "on" Then 
		strProtocol = "https" 
	Else
		strProtocol = "http" 
	End If

	strDomain= Request.ServerVariables("SERVER_NAME")
	strPath= Request.ServerVariables("SCRIPT_NAME") 
	strQueryString= Request.ServerVariables("QUERY_STRING")

	strFullUrl = strProtocol & "://" & strDomain & strPath
	If Len(strQueryString) > 0 Then
	   strFullUrl = strFullUrl & "?" & strQueryString
	End If
	GetCurrentUrl = strFullUrl
End Function
%>