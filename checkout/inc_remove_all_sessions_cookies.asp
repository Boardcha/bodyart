<%
' Remove all session variables

	'response.write "# of session variables: " & Session.Contents.Count
	for each item in session.contents
		if item <> "sandbox" and item <> "custID_account" and item <> "sandbox_toggle" and item <> "invoiceid" and item <> "cc_status" and item <> "cim_accountNumber" and item <> "google_sent" then
			'response.write "Session variable: " & item & " value: " & session(item) & "<br/>"
			session(item) = ""
		end if	
	next
'	Session.Abandon

	
' Remove all cookie variables except ID
For Each x in Request.Cookies
'response.write "-----" & x & "-----" & request.Cookies(x) & "-----<br>"
	' Do not delete these cookies below (login cookies)
	if (request.cookies(x) <> request.cookies("ID")) and (request.cookies(x) <> request.cookies("token")) and (request.cookies(x) <> request.cookies("selector")) and (request.cookies(x) <> request.cookies("cartSessionid")) and (request.cookies(x) <> request.cookies("cartSelector")) and (request.cookies(x) <> request.cookies("CookieInfoScript")) then
	
		Response.Cookies(x) = ""
		Response.Cookies(x).Expires = DateAdd("d",-1,now())

	end if
Next

' Set cart count to 0
	response.cookies("cartCount") = 0
%>
<script type="text/javascript">
	// Remove all local storage variables
	localStorage.clear();
</script>