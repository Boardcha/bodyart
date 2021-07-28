<% @LANGUAGE="VBSCRIPT" %>
<!--#include virtual="/template/inc_includes.asp" -->
<!--#include virtual="functions/token.asp"-->
<!--#include virtual="functions/hash_extra_key.asp"-->

<% if Request.form("email") <> "" then 

	' clear out tokens that are older than 24 hours
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE customers SET reset_token = '' WHERE reset_token_timestamp <= '" & now()-2 & "'"
	objCmd.Execute()

	' REQUEST A RESET TOKEN -------------------------------------
	if request.form("email") <> "" then
	
		mailer_type = "front_reset_user_password"
		reset_token = getToken(40, extraChars)
		hashed_token = sha256(reset_token)
	
		' DOES TOKEN EXIST? ----------------------------------------
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT reset_token FROM customers WHERE reset_token = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("token",200,1,250,hashed_token))
		Set rsTokenExists = objCmd.Execute()
	
		if rsTokenExists.BOF and rsTokenExists.EOF then ' if end of file (token doesn't exist)
		
			'Check to see if user account exists
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "SELECT email FROM customers WHERE email = ?"
			objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,250,Request.Form("email")))
			Set rsAccountFound = objCmd.Execute()
			
			if NOT rsAccountFound.BOF and NOT rsAccountFound.EOF then ' if account is found
		
				var_show_success_message = "yes"

				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE customers SET reset_token = '" & sha256(reset_token) & "', reset_token_timestamp = '" & now() & "' WHERE email = ?"
				objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,250,Request.Form("email")))
				objCmd.Execute()
				
%>
	<!--#include virtual="emails/function-send-email.asp"-->
	<!--#include virtual="emails/email_variables.asp"-->
<%
			
			else ' if account not found
			
				var_account_not_found = "yes"
			
			end if ' if account found/not found

			set rsAccountFound = nothing
		end if ' if rsTokenExists is not empty
		set rsTokenExists = Nothing
		
	end if ' if email found
	' REQUEST A RESET TOKEN -------------------------------------

end if 



DataConn.Close()
Set DataConn = Nothing
%>


<% if var_account_not_found = "yes" then %>
	<div class="alert alert-danger">
		<strong>No account found</strong>
		<br/><br/>
		<a class="alert-link" data-toggle="modal" data-target="#createaccount"
		data-dismiss="modal" href="">Click here</a> to create a new account
	</div>
<% end if %>
<% if var_show_success_message = "yes" then %>
	<div class="alert alert-success">
		Your password reset link has been sent to the e-mail you provided. It should arrive in your inbox within a few minutes. If you don't see it, please check your junk mail folder. The e-mail link sent will be valid for 24 hours to update your password.<br/>
		<br/>
		If you need more assistance accessing your account feel free to contact our <a class="alert-link"  href="/contact.asp">customer service department</a> and we'll be happy to help!
	</div>
<% end if %>