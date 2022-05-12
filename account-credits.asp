<% @LANGUAGE="VBSCRIPT" %>
<%
	page_title = "Your account profile"
	page_description = "Your Bodyartforms account. Edit your profile, view orders, and more."
	page_keywords = ""
	
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/cart/inc_cart_main.asp"-->
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

' Pull gift certificates
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT custID, amount, code FROM dbo.TBLcredits WHERE amount <> 0 AND custID = ? ORDER BY ID DESC"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
		Set rsGetCertificates = objCmd.Execute()

' Pull customer photo gallery
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT ProductID, filename, thumb_filename, PhotoID, customerID, description, title, gauge, length, ProductDetail1, DateSubmitted FROM dbo.QRY_Photos WHERE customerID = ? ORDER BY PhotoID DESC"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
		Set rsGetPhotos = objCmd.Execute()
		
		' Get address book
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM TBL_Addressbook WHERE custID = ? AND address_type = 'shipping' ORDER BY default_shipping DESC"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
		Set rsGetShippingAddresses = objCmd.Execute()
		
		' Get billing addresses
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM TBL_Addressbook WHERE custID = ? AND address_type = 'billing' ORDER BY default_billing DESC"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
		Set rsGetBillingAddresses = objCmd.Execute()

		'Get country list for drop downs
		Set rsGetCountrySelect = Server.CreateObject("ADODB.Recordset")
		rsGetCountrySelect.ActiveConnection = DataConn
		rsGetCountrySelect.Source = "SELECT * FROM dbo.TBL_Countries WHERE Display = 1 ORDER BY Country ASC "
		rsGetCountrySelect.CursorLocation = 3 'adUseClient
		rsGetCountrySelect.LockType = 1 'Read-only records
		rsGetCountrySelect.Open()
		
		
End if ' Only run this info if a match was found

%>
<div class="display-5">
		Account Credits &amp; Discounts
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
<% elseif rsGetUser("active") = 0 then %>
	<div class="alert alert-danger"><h5>Your account has not been activated yet.</h5>Please click on the activation link sent to your email to confirm your account registration and access your account.</div>
<% else %>

<%
' -----------  LAUNCHED GRANDFATHERED DATE ON 1/7/2019 -----------
if rsGetUser.Fields.Item("grandfathered_discount").Value = 1 then %>
<div class="card mt-4">
	<div class="h5 card-header">
		Customer loyalty discount status
	</div>
	<div class="card-body">
		<% If TotalSpent > 275 Then %>
			<div class="alert alert-success">
				You're getting 10% off every order you place
			</div>
		<% else %>
			<div class="alert alert-secondary">
				<span class="bold">$<%= 275 - TotalSpent %></span> to go before you can qualify for 10% OFF EVERY FUTURE ORDER (Spend $275 or more after shipping &amp; discounts. Discounts are not retro-active.)
			</div>
		<% end if %>
	</div>
  </div>
<% end if %>

		<% if (rsGetUser.Fields.Item("credits").Value) <> 0 then 
			var_hide_credits = " "
		else
			var_hide_credits = " style='display:none'"
		end if
		%>

		<div class="card mt-4" <%= var_hide_credits %>>
			<div class="h5 card-header">
				Available store credit
			</div>
			<div class="card-body">
				Store credit: <span class="alert alert-success p-1 m-0 current-credit"><%= FormatCurrency(rsGetUser.Fields.Item("credits").Value, 2)%></span>
			</div>
		  </div>

		
			
		<% if rsGetUser.Fields.Item("Points").Value > 0 then %>
		<div class="card mt-4 block-redeem-points">
			<div class="h5 card-header">
				Your points
			</div>
			<div class="card-body">
				<button class="btn btn-purple btn confirm-redeem" type="button">Redeem points</button>
        
				<div class="py-3">You currently have <span class="alert alert-success p-1 m-0"><%= rsGetUser.Fields.Item("Points").Value %> points worth <%= FormatCurrency(rsGetUser.Fields.Item("Points").Value * .05, 2)%></span> in store credit to redeem</div>    
            
        		<div class="small">* Points are earned for reviewing jewelry (1 point) or submitting photos of you wearing our jewelry (3 points).</div>
			</div>
		  </div>
		<% end if 'rsGetUser.Fields.Item("Points").Value > 0
        %>
        
		<div class="message-redeem-points"></div>
		


		<div class="card mt-4">
			<div class="h5 card-header">
				Convert a gift certificate to store credit
			</div>
			<div class="card-body">
				<div class="form-inline">
					<label for="cert-code"><i class="fa fa-gift fa-2x text-secondary mt-2"></i></label><input class="form-control mx-2 mt-2 w-auto" type="text" name="cert-code" id="cert-code" placeholder="Enter gift code here" style="display:inline-block">
					<button class="btn btn-purple mt-2 btn-convert-cert" type="button">Add</button>
					</div>
			
					<div class="message-convert-cert"></div>
					<% If Not rsGetCertificates.eof then %>
					<br/>
					<h6>Your gift certificates</h6>
				   
					<% Do While Not rsGetCertificates.EOF %>
					<%= FormatCurrency(rsGetCertificates.Fields.Item("amount").Value,2)%>&nbsp;&nbsp;&nbsp;<%=(rsGetCertificates.Fields.Item("code").Value)%><br />                         
					<% 
					rsGetCertificates.MoveNext()
					Loop
					
					end if ' Not rsGetCertificates.eof
					%>
			</div>
		  </div>		
    
	
	
	
<%
end if   'rsGetUser.EOF
%>  


<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript">

	// START redeem points
	$('.confirm-redeem').click(function () {
		$.ajax({
		method: "post",
		dataType: "json",
		url: "accounts/ajax-redeem-points.asp"
		})
		.done(function(json, msg) {
			$('.message-redeem-points').html('<div class="alert alert-success alert-dismissible fade show" role="alert">Your points have been redeemed and your store credit has been updated with the new amount.  <button type="button" class="close" data-dismiss="alert" aria-label="Close"><span aria-hidden="true">&times;</span></button></div>');
			$('.block-redeem-points').delay(5000).fadeOut('slow');
			$('.section-credits').show();
			$('.current-credit').html(json.amount);
		})
		.fail(function(json, msg) {
			$('.message-redeem-points').html('<div class="alert alert-danger">Site error. Please contact customer service for assistance.</div>');
		})
	});  // END redeem points

	// START convert gift cert to store credit
	$('.btn-convert-cert').click(function () {
		var cert_code = $('#cert-code').val();
		
		$.ajax({
		method: "post",
		dataType: "json",
		url: "accounts/ajax-convert-gift-cert.asp",
		data: {cert_code: cert_code}
		})
		.done(function(json, msg) {
			if (json.status == 'success') {
				$('.message-convert-cert').html('<div class="alert alert-success">Your gift certificate has been added to your store credit.</div>').show();
				$('.message-convert-cert').delay(5000).fadeOut('slow');
				$('.section-credits').show();
                $('.current-credit').html(json.amount);
                $('#cert-code').val('');
			}
			if (json.status == 'fail') {
				$('.message-convert-cert').html('<div class="alert alert-danger">Gift certificate code entered was not found. Please double check your code or contact customer service for assistance.</div>').show();
			}
		})
		.fail(function(json, msg) {
			$('.message-convert-cert').html('<div class="alert alert-danger">Site error. Please contact customer service for assistance.</div>').show();
		})
	});  // END convert gift cert to store credit
</script>