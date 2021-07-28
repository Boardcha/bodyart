<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.querystring("token") <> "" then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT ID FROM TBL_AdminUsers WHERE reset_token = ? AND reset_token_timestamp >= '" & now() - 1 & "'"
	objCmd.Parameters.Append(objCmd.CreateParameter("token",200,1,150,sha256(request.querystring("token"))))
	Set rs_getUser = objCmd.Execute()
	
	if NOT rs_getUser.BOF and NOT rs_getUser.EOF then
		
		' if user is found, and update requested, then update
		
	else
	'	response.redirect "password_reset.asp?token=" & request.querystring("token") ' if token not found
	end if
	
end if ' if form fields are not empty
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../CSS/Admin.css" />
<!--#include file="includes/inc_scripts.asp"-->
<script type="text/javascript" src="../js/jquery.form-validator.min.js"></script>
<script type="text/javascript" src="../js/security.js"></script>
<script type="text/javascript" src="../js/password_generator.jquery.js"></script>
<script type="text/javascript">
$(document).ready(function() {
	$('#password-input').removeAttr('title');

	// Update password
	$("form").submit(function(){
		var password = $('#pass').val();
		var token = $('#pass').attr("data-token");
		
		$.ajax({
		method: "POST",
		dataType: "json",
		url: "administrative/ajax_reset_user_password.asp",
		data: {token: token, password: password}
		})
		.done(function( json, msg ) {
			console.log("Success");
			if (json.status == "success") {
				$('#frm_update').hide();
				$('#update-success').show();
			}
			if (json.status == "fail") {
				window.location.replace("?token=" + token);
			}
		})
		.fail(function(json, msg) {
			console.log("ajax failed");
		});
	//	e.preventDefault(); // don't refresh after submit
	return false;
	});
	
	
	$.validate({
		modules : 'security',
		onModulesLoaded : function() {
		var optionalConfig = {
		fontSize: '9pt',
		padding: '4px',
		bad : 'Weak password',
		weak : 'Weak password',
		good : 'Decent password',
		strong : 'Strong password'
		};

		$('#password-input').displayPasswordStrength(optionalConfig);
		}
	});
	
  
  $('#password-generator-button').pGenerator({
    'bind': 'click',
    'passwordElement': '#password-input, #pass',
 //   'displayElement': '#display-password',
    'passwordLength': 12,
    'uppercase': true,
    'lowercase': true,
    'numbers':   true,
    'specialChars': true,
    'onPasswordGenerated': function(generatedPassword) {
		$('#password-input').val(generatedPassword).trigger('change');
		$('#password-input, #pass').removeClass("error"); 
		$('#password-input, #pass').removeAttr('style');
		$('.help-block, .strength-meter').hide();
		}
	});

});  
</script>
<title>Admin login</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="content-grey">
	<div class="section-headers">
		Reset password
	<hr class="hr">
	</div> 
	<br/>
	<div id="update-success" class="no-display notice-eco">
		The password has been updated. <a href="login.asp">Click here</a> to login.
	</div>
	<form id="frm_update" class="admin-fields" style="padding-left: 50px">
<% if NOT rs_getUser.BOF and NOT rs_getUser.EOF then ' if user found by token 

'  pattern="(?=.*\d)(?=.*[a-z])(?=.*[A-Z]).{10,}"
%>

			<div style="padding: 5px 5px 5px 0;">Password (Minimum 10 characters)
				&nbsp;&nbsp;&nbsp;<a href="" id="password-generator-button">Generate password</a>
				&nbsp;&nbsp;&nbsp;<span id="display-password"></span>
			</div>
			<div class="control-group">
			<input type="text" name="pass_confirmation" id="password-input" autofocus = "autofocus"  autocomplete="off" required  minlength=10 pattern="(?=^.{10,}$)((?=.*\d)(?=.*\W+))(?![.\n])(?=.*[A-Z])(?=.*[a-z]).*$" title="Must be at least 10 or more characters. Must contain at least one of each: number, uppercase letter,  lowercase letter, and special character." oninvalid="this.setCustomValidity('Password does not meet requirements: \nMust be at least 10 or more characters. Must contain at least one of each: number, uppercase letter,  lowercase letter, and special character.')" oninput="setCustomValidity('')" data-validation="strength" data-validation-strength="3" />
			</div>
			
		<br/>
			<div style="padding: 5px 5px 5px 0;">Re-type password</div>
			<div class="control-group">
				<input type="text" name="pass" id="pass" autocomplete="off" data-validation="confirmation" data-validation-error-msg="Passwords don't match" data-token="<%= request.querystring("token") %>" />
			</div>
		<br/>
			<button type="submit" formmethod="post" formaction="">Update password</button>
		<br/>
		<br/>
		<br/>
		<strong>WRITE PASSWORD DOWN</strong>
		<br/>
		Once the password has been updated, it can not be retrieved from the database for security reasons.
<br/>
<% else ' if user by token not found
%>
	<div class="notice-red">
		Sorry, but the reset token has expired or is not found. Please request a new one.
	</div>
<% end if ' if user not found %>	
	</form>
</div>
</body>
</html>
<%
	Set rs_getUser = Nothing
	DataConn.Close()
	Set DataConn = Nothing
%>