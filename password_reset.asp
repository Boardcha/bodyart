<% @LANGUAGE="VBSCRIPT" %>
<%
	page_title = "Bodyartforms password reset"
	page_description = "Reset your Bodyartforms account password"
	page_keywords = ""
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<!--#include virtual="functions/token.asp"-->
<!--#include virtual="functions/hash_extra_key.asp"-->
<%
if request.querystring("token") <> "" then

	salt = getSalt(32, extraChars)
	newPass = sha256(salt & request.form("password") & extra_key)
	token = sha256(request.form("token"))

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT customer_ID FROM customers WHERE reset_token = ? AND reset_token_timestamp >= '" & now() - 1 & "' ORDER BY customer_ID ASC"
	objCmd.Parameters.Append(objCmd.CreateParameter("token",200,1,150,sha256(request.querystring("token"))))
	Set rsTokenExists = objCmd.Execute()
	
	if NOT rsTokenExists.BOF and NOT rsTokenExists.EOF then
		
		' if user is found, and update requested, then update
		if request.form("password") <> "" then

			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "UPDATE customers SET pass_last_updated = '" & now() & "',  password_hashed = ?, salt = ?, reset_token = '', reset_token_timestamp = '' WHERE customer_ID = ?" 
			objCmd.Parameters.Append(objCmd.CreateParameter("password",200,1,250,newPass))
			objCmd.Parameters.Append(objCmd.CreateParameter("salt",200,1,250,salt))
			objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,rsTokenExists.Fields.Item("customer_ID").Value))
			objCmd.Execute()
			
			var_pass_updated = "yes"
			
		end if ' if password field is not empty
		
	else
	'	response.redirect "password_reset.asp?token=" & request.querystring("token") ' if token not found
	end if
	
end if ' if form fields are not empty
%>


<div class="display-5 mb-5">
	Reset Your Password
</div>
	<div id="update-success" class="alert alert-success" <% if var_pass_updated <> "yes" then %>style="display:none"<% end if %>>
		The password has been updated. <a class="alert-link" href="#" data-toggle="modal" data-target="#signin">Click here</a> to login.
	</div>

<%
if var_pass_updated <> "yes" then
%>
	<form class="needs-validation" id="frm-reset-password" novalidate>
<% if NOT rsTokenExists.BOF and NOT rsTokenExists.EOF then ' if user found by token 

%>
<div class="form-group">
	<label for="password_change">Password (Minimum 8 characters)</label>
	<input class="form-control w-50" type="password" name="password" id="password_change" autocomplete="off" minlength=8 required>
	<div class="invalid-feedback">
		Password is required (Minimum 8 characters)
</div>
</div>
<div class="form-group">
	<label for="password_change_confirmation">Re-type password</label>
	<input class="form-control w-50" type="password" name="password-confirmation" id="password_change_confirmation" data-validation="confirmation"  autocomplete="off" minlength=8 required>
	<div class="invalid-feedback">
			Password is required
</div>
<div class="text-danger small" id="nomatch"></div>
</div>

	<button class="btn btn-purple" type="submit" id="btn-update-pass" formmethod="post" formaction="password_reset.asp?token=<%= request.querystring("token") %>" >Update password</button>
<br/>
<% else ' if user by token not found
%>
	<div class="alert alert-danger">
		Sorry, but the reset token has expired or is not found. Please request a new one.
	</div>
<% end if ' if user not found 
set rsTokenExists = nothing
%>	
	</form>
<% end if ' if pass has not been submitted to update yet
%>



<!--#include virtual="/bootstrap-template/footer.asp" -->
<script>

	// Compare passwords
	$("#password_change_confirmation").blur(function(){
		var password = $("#password_change").val();
    	var confirmPassword = $("#password_change_confirmation").val();
				 // Check for equality with the password inputs
				 if (password != confirmPassword ) {
                    $('#nomatch').html('Passwords do not match');
                    $('#btn-update-pass').prop('disabled', true);
                } else {
                    $('#nomatch').html('');
                    $('#btn-update-pass').prop('disabled', false);
                }        
    });


	$('#frm-reset-password').submit(function (e) {
// Fetch form to apply custom Bootstrap validation
var form = $("#frm-reset-password")

	if (form[0].checkValidity() === false) {
		e.preventDefault()
		e.stopPropagation()
		console.log("invalid form elements");
	
		form[0].classList.add('was-validated');
		e.preventDefault();
	}
        
});
</script>