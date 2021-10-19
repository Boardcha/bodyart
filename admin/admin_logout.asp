<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' *** Logout the current user.
MM_logoutRedirectPage = "/admin/login.asp?login=yes"
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

If (MM_logoutRedirectPage <> "") Then Response.Redirect(MM_logoutRedirectPage)
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<title>Admin logout</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body  topmargin="0" class="mainbkgd" >
  <!--#include file="admin_header.asp"-->
<span class="adminheader">Admin logout </span>
</body>
</html>