<% @LANGUAGE="VBSCRIPT" %>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<%
email = Request("email")
activation_hash = Request("hash")

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM customers WHERE email = ? AND activation_hash = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,250,email))
objCmd.Parameters.Append(objCmd.CreateParameter("activation_hash",200,1,250,activation_hash))
Set rsGetUser = objCmd.Execute()
		
If Not rsGetUser.EOF Then
	Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
	password = "3uBRUbrat77V"
	data = rsGetUser.Fields.Item("customer_ID").Value
	encrypted = objCrypt.Encrypt(password, data)
	Response.Cookies("ID") = encrypted
	Response.Cookies("ID").Expires = DATE + 60
	Set objCrypt = Nothing	
End If	
%>
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<div class="display-5 mb-3">
	Account Activation
</div>
<% If rsGetUser.EOF Then%>
	<div class="alert alert-danger">Invalid activation code. Please click on the link in the email sent upon registration.</div>
<%Else%>
	<%
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET active = 1 WHERE email = ? AND activation_hash = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,250,email))
	objCmd.Parameters.Append(objCmd.CreateParameter("activation_hash",200,1,250,activation_hash))
	objCmd.Execute()
	%>	
	<div class="alert alert-success">Your account is now active! Go to your <a href="account.asp">Account</a> page.</div>
<%End If%>
<%
If Not rsGetUser.EOF Then
	mailer_type = "new account"
	'Send mail out only when the user used the hash first time
	If rsGetUser("active") = false Then%>
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/emails/email_variables.asp"-->	
	<%End If
End If%>	
<!--#include virtual="/bootstrap-template/footer.asp" -->