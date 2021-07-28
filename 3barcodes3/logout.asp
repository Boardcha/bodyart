<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' *** Logout the current user.
MM_logoutRedirectPage = "index.asp"
Session.Contents.Remove("MM_Username")
Session.Contents.Remove("MM_UserAuthorization")


'Delete any authentication tokens for user
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "DELETE FROM tbl_admin_users_auth_tokens WHERE user_id = " & user_id & " AND user_agent='" & Left(Request.ServerVariables("HTTP_USER_AGENT"), 300) & "'"
objCmd.Execute()	

Response.Cookies("token") = ""
Response.Cookies("token").Expires = DATE-1
Response.Cookies("selector") = ""
Response.Cookies("selector").Expires = DATE-1
Session.Contents.Remove("SubAccess")

if user_id = "" then 
    user_id = 0
end if
%>
<html>
<head>
<link href="/CSS/baf.min.css?v=092519" rel="stylesheet" type="text/css" />
<title>Scanners logout</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0" />
<meta name="mobile-web-app-capable" content="yes">
</head>
<body class="p-3">
        <a class="btn btn-primary" href="login.asp" role="button">Login again</a>

</body>
</html>