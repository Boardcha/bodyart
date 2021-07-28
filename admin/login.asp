<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../functions/salt.asp"-->
<!--#include file="../functions/hash_extra_key.asp"-->
<!--#include file="../functions/token.asp"-->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

if request.form("username") <> "" and request.form("password") <> "" then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT username, password_hashed, salt, ID FROM TBL_AdminUsers WHERE username = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("username",200,1,50,Request.form("username")))
	Set rs_getUser = objCmd.Execute()
	
	if NOT rs_getUser.BOF and NOT rs_getUser.EOF then
	
		var_user_id = rs_getUser.Fields.Item("ID").Value
		
		While NOT rs_getUser.EOF
		
			hashed_password = sha256(rs_getUser.Fields.Item("salt").Value & request.form("password") & extra_key)
			cookie_token = getToken(64, extraChars)
			cookie_selector = Left(sha256(getToken(12, extraChars)),12)
					
			if hashed_password = rs_getUser.Fields.Item("password_hashed").Value then
									
				'Write authentication tokens to database to keep user logged in via cookie
			
				'Delete any authentication tokens for user
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "DELETE FROM tbl_admin_users_auth_tokens WHERE user_id = " & var_user_id & " AND user_agent='" & Left(Request.ServerVariables("HTTP_USER_AGENT"), 300) & "'"
				objCmd.Execute()						
				
				response.cookies("token") = cookie_token
				response.cookies("token").Expires = DATE + 7
				response.cookies("selector") = cookie_selector
				response.cookies("selector").Expires = DATE + 7
				
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "INSERT INTO tbl_admin_users_auth_tokens (expires, token, selector, user_id, user_agent) VALUES ('" & now()+3 & "', ?, ?, " & var_user_id & ", ?)"
				objCmd.Parameters.Append(objCmd.CreateParameter("cookie_token",200,1,250,sha256(cookie_token & cookie_extra_key)))
				objCmd.Parameters.Append(objCmd.CreateParameter("selector",200,1,250,cookie_selector))
				objCmd.Parameters.Append(objCmd.CreateParameter("user_agent",200,1,250, Left(Request.ServerVariables("HTTP_USER_AGENT"), 300)))
				objCmd.Execute()
				
				response.redirect "index.asp"
			else
				response.redirect "login.asp?status=failed&login=yes" ' if pass incorrect
			end if
		
		rs_getUser.MoveNext()
		Wend
		
	else
		response.redirect "login.asp?status=failed&login=yes" ' if user/pass not found
	end if
	
	DataConn.Close()
	Set DataConn = Nothing

end if ' if form fields are not empty
%>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--#include file="includes/inc_scripts.asp"-->
<title>Admin login</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
	<h5>
		Admin login
	</h5> 
	<form class="mt-4">
		<div class="form-group">
			<label for="username">Username</label>
			<input class="form-control form-control-sm w-25" type="text" name="username">
		</div>
		<div class="form-group">
			<label for="password">Password</label>
			<input class="form-control form-control-sm w-25" type="password" name="password">
		</div>

			<button class="btn btn-purple" type="submit" formmethod="post" formaction="login.asp">Login</button>

<% if request.querystring("status") = "failed" then %>
	<div class="alert alert-red">
		Sorry, but the username or password does not match. Please try again.
	</div>
<% end if %>	
	</form>
</div>
</body>
</html>