<% @LANGUAGE="VBSCRIPT" %>
<%
	page_title = "Your account profile"
	page_description = "Your Bodyartforms account. Edit your profile, view orders, and more."
	page_keywords = ""
	
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->

<!--#include virtual="/bootstrap-template/filters.asp" -->
<%
var_flagged = ""

' Pull the customer information from a cookie
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM customers  WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
		Set rsGetUser = objCmd.Execute()
		
If Not rsGetUser.EOF Or Not rsGetUser.BOF Then ' Only run this info if a match was found



	if rsGetUser.Fields.Item("Flagged").Value = "Y" then
		var_flagged = "yes"
	end if
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT shipped, email FROM sent_items WHERE email = ? AND (shipped = 'Flagged' OR shipped = 'Chargeback')"
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,250,rsGetUser.Fields.Item("email").Value))
	set rsGetFlaggedOrders = objCmd.Execute()
	
	if NOT rsGetFlaggedOrders.eof then
		var_flagged = "yes"
	end if		
		
End if ' Only run this info if a match was found

%>

<div class="display-5">
		Account Profile
	</div>


<%
if session("admin_tempcustid") <> "" then %>
	<div class="alert alert-success">Admin viewing
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<a href="account.asp?cleartemp=yes">Reset</a>
	</div>
<% end if %>
<!--#include virtual="/accounts/inc-account-navigation.asp" -->
<% If rsGetUser.EOF or var_flagged = "yes" Then
%>
	<div class="alert alert-danger">Not logged in or no account found</div>
<% else %>

	<% ' variable set on signin_transfer.asp -- Shows notice to user if NON logged in cart items got transferred to their account cart
	if session("cart_items_transferred") = "yes" then
	%>
	<div class="alert alert-success">
		All items have been transferred to your account cart.
	</div>
	<%
		session("cart_items_transferred") = ""
	end if
	%>
	
<div class="card card-light mt-4">
	<div class="card-header">
		<h5>Update your profile name</h5>
	</div>
	<div class="card-body">
		<form class="needs-validation col-md-4 p-0" name="frm-update-profile" id="frm-update-profile" novalidate>
			<div class="form-group">
			<label for="updatefirst">First name <span class="text-danger">*</span></label>
			<input class="form-control" name="first" type="text" id="updatefirst" value="<%= rsGetUser.Fields.Item("customer_first").Value %>" required />
			<div class="invalid-feedback">
					First name is required
			</div>
		</div>
		<div class="form-group">
			<label for="updateLast">Last name <span class="text-danger">*</span></label>
			<input class="form-control" name="last" id="updateLast" type="text" value="<%= rsGetUser.Fields.Item("customer_last").Value %>"  required />
			<div class="invalid-feedback">
					Last name is required
			</div>
			</div>
			<button type="submit" class="btn btn-purple">Update name</button> 
		</form>
		<div class="message-update-profile"></div>
	</div><!-- end card body -->
</div><!-- end card -->	


<div class="card card-light mt-5">
		<div class="card-header">
				<h5>Update your e-mail address</h5>
		</div>
		<div class="card-body">
		<form class="needs-validation col-md-4 p-0" name="frm-update-email" id="frm-update-email" novalidate>
			<div class="form-group">
					<label for="updateEmail">E-mail <span class="text-danger">*</span></label>
			<input class="form-control" name="email" type="email" id="updateEmail" value="<%= rsGetUser.Fields.Item("email").Value %>" required />
			<div class="invalid-feedback">
					A valid e-mail is required
			</div>
			</div>
			<button type="submit" class="btn btn-purple" id="btn-update-email">Update e-mail</button> 
		</form>
		<div class="message-update-email"></div>
	</div><!-- end card body -->
</div><!-- end card -->	


<div class="card card-light mt-5">
		<div class="card-header">
				<h5>Update your password</h5>
		</div>
		<div class="card-body">
	<form class="needs-validation col-md-4 p-0" name="frm-update-pass" id="frm-update-pass" novalidate>
		<%If rsGetUser("registered_with_social_login").Value <> True Then%>
		<div class="form-group">
			<label for="current_password">Current password: <span class="text-danger">*</span></label>
			<input class="form-control" name="current_password" id="current_password" type="password"  required/>
			<div class="invalid-feedback message-check-pass">
					Current password is required
			</div>
		</div>
		<%End If%>	
		<div class="form-group">
			<label for="password_confirmation">New password: <span class="text-danger">*</span></label>
			<input class="form-control" name="password_confirmation" id="profile_password_confirmation" type="password" required/>
			<div class="invalid-feedback">
					New password is required
			</div>
		</div>
		<div class="form-group">
			<label for="password">Re-type pasword: <span class="text-danger">*</span></label>
			<input class="form-control" name="password" type="password" id="first_password" id required />
			<div class="invalid-feedback">
					Password confirmation is required
			</div>
		</div>
		<input type="hidden" name="customer_id" id="customer_id" value="<%= rsGetUser.Fields.Item("customer_ID").Value %>" />
		<button type="submit" class="btn btn-purple" id="btn-update-pass">Update password</button>
	</form>
		<div class="message-update-pass"></div>
	</div><!-- end card body -->
</div><!-- end card -->	
<%
end if   'rsGetUser.EOF
%>  


<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript" src="/js-pages/account-profile.min.js?v=113019"></script> 