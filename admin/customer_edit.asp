<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/authnet.asp"-->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


' Get customer info
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT customer_ID, customer_first, customer_last, email, credits, moderator, credits_Contests, Flagged, FlagNotes, points, account_created, cim_custid FROM dbo.customers WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,Request.QueryString("ID")))
Set rsGetCustomer = objCmd.Execute()

' Get SHIPPING address book
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM TBL_Addressbook WHERE custID = ? AND address_type = 'shipping' ORDER BY default_shipping DESC"
objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,Request.QueryString("ID")))
Set rsGetShippingAddresses = objCmd.Execute()

' Get BILLING profiles
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM TBL_Addressbook WHERE custID = ? AND address_type = 'billing' ORDER BY default_billing DESC"
objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,Request.QueryString("ID")))
Set rsGetBillingAddresses = objCmd.Execute()

%>
<html>
<head>
<title>Edit customer info</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h4 class="d-inline">
		Customer profile
		<h6 class="ml-5 d-inline small small">Created <%=(rsGetCustomer.Fields.Item("account_created").Value)%></h6>
</h4> 

<form class="ajax-update">

<table class="table table-sm table-striped mt-4">
	<thead class="thead-dark">
		<tr>
			<th>First</th>
			<th>Last</th>
			<th>Email</th>
			<th>Store credit</th>
			<th>Points</th>
			<th>Moderator</th>
			<th>Flagged</th>
		</tr>
	</thead>
	<tr>
		<td>
			<input class="form-control form-control-sm" type="text" name="fname" size="10" value="<%=(rsGetCustomer.Fields.Item("customer_first").Value)%>" data-id="<%= rsGetCustomer.Fields.Item("customer_ID").Value %>" data-column="customer_first" data-friendly="First name" data-int_string="string">
		</td>
		<td>
			<input class="form-control form-control-sm" type="text" name="lname" size="10" value="<%=(rsGetCustomer.Fields.Item("customer_last").Value)%>" data-id="<%= rsGetCustomer.Fields.Item("customer_ID").Value %>" data-column="customer_last" data-friendly="Last name" data-int_string="string">
		</td>
		<td>
			<input class="form-control form-control-sm" name="email" type="text" id="email" value="<%=(rsGetCustomer.Fields.Item("email").Value)%>" size="30" data-id="<%= rsGetCustomer.Fields.Item("customer_ID").Value %>" data-column="email" data-friendly="Email" data-int_string="string">
		</td>
		<td class="form-inline">
			$
			<input class="form-control form-control-sm ml-2" name="varcredit" type="text" value="<%= FormatNumber(rsGetCustomer.Fields.Item("credits").Value, -1, -2, -2, -2) %>" size="5" data-id="<%= rsGetCustomer.Fields.Item("customer_ID").Value %>" data-column="credits" data-friendly="Store credit" data-int_string="money">
		</td>
		<td>
			<input class="form-control form-control-sm" name="points" type="text" id="points" value="<%=(rsGetCustomer.Fields.Item("points").Value)%>" size="5" maxlength="5" data-id="<%= rsGetCustomer.Fields.Item("customer_ID").Value %>" data-column="points" data-friendly="Points" data-int_string="integer">
		</td>
		<td>
			<input type="radio" name="moderator" id="moderator" value="yes" <% if (rsGetCustomer.Fields.Item("moderator").Value) = "yes" then %>checked<% end if %> data-id="<%= rsGetCustomer.Fields.Item("customer_ID").Value %>" data-column="moderator" data-friendly="Moderator" data-int_string="string">
			  Yes
			  <br/><br/>
			  <input name="moderator" type="radio" id="moderator" value="no" <% if (rsGetCustomer.Fields.Item("moderator").Value) <> "yes" then %>checked<% end if %> data-id="<%= rsGetCustomer.Fields.Item("customer_ID").Value %>" data-column="moderator" data-friendly="Moderator" data-int_string="string">
			  No
		</td>
		<td>
			<input name="Flagged" type="checkbox" id="Flagged" value="Y" <% if (rsGetCustomer.Fields.Item("Flagged").Value) = "Y" then %>checked<% else %><% end if %> data-id="<%= rsGetCustomer.Fields.Item("customer_ID").Value %>" data-column="Flagged" data-friendly="Flagged" data-int_string="string">
			<textarea class="form-control form-control-sm mt-2" name="FlagNotes" cols="40" rows="3" id="FlagNotes" data-id="<%= rsGetCustomer.Fields.Item("customer_ID").Value %>" data-column="FlagNotes" data-friendly="Account notes" data-int_string="string" placeholder="Notes on why account is flagged"><%= rsGetCustomer.Fields.Item("FlagNotes").Value %></textarea>
		</td>
	</tr>
</table>
    
</form>


<%
	if not rsGetBillingAddresses.eof then
%>
	<table class="table table-sm table-hover table-striped mt-4">
		<thead class="thead-dark">
			<tr>
			<th colspan="4">
				<h5>Saved credit cards</h5>
				<div class="form-inline">
				$
				<input class="form-control form-control-sm mx-2" type="text" name="charge-cim" size="5" id="charge-cim">
				<span>Amount to charge to card</span>
				</div>
			</th></tr>
			<tr>
				<th>Card #</th>
				<th>Address</th>
				<th>Added</th>
				<th>Last Updated</th>
			</tr>
		</thead>
<%
			while NOT rsGetBillingAddresses.EOF
			
			' Connect to Authorize.net CIM to get CREDIT CARD information
			strGetBillingAddress = "<?xml version=""1.0"" encoding=""utf-8""?>" _
			& "<getCustomerPaymentProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
			& MerchantAuthentication() _
			& "  <customerProfileId>" & rsGetCustomer.Fields.Item("cim_custid").Value & "</customerProfileId>" _
			& "  <customerPaymentProfileId>" & rsGetBillingAddresses.Fields.Item("cim_shippingid").Value & "</customerPaymentProfileId>" _
			& "</getCustomerPaymentProfileRequest>"
			
			Set objResponseGetAddress = SendApiRequest(strGetBillingAddress)

			' If connection is a success to address book than retrieve values and assign to variables
			If IsApiResponseSuccess(objResponseGetAddress) Then
				strBilling_cardnumber = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:payment/api:creditCard/api:cardNumber").Text	


				strBilling_first = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:firstName").Text
				strBilling_last = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:lastName").Text
				'Split out state from authorize.net and break it out to address 1 and address 2 fields
				split_address_billing = Split(objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:address").Text, "|")
					strBilling_address = split_address_billing(0)
					strBilling_address2 = split_address_billing(1)
						
				strBilling_city = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:city").Text
				strBilling_state = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:state").Text
				strBilling_zip = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:zip").Text
				strBilling_country = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:country").Text
				strBilling_ID = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:customerPaymentProfileId").Text
			Else
			'	Response.Write "The operation failed with the following errors:<br>" & vbCrLf
			'	PrintErrors(objResponseGetAddress)
			End if
			
		%>
		<tr>
			<td>
				<span id="loading-<%= rsGetBillingAddresses.Fields.Item("cim_shippingid").Value %>"></span>
				<span class="button_small_grey charge-button" style="display:none" data-cim_account_id="<%= rsGetCustomer.Fields.Item("cim_custid").Value %>" data-cim_billing_id="<%= rsGetBillingAddresses.Fields.Item("cim_shippingid").Value %>">Charge this card $<span class="charge-amount"></span></span>&nbsp;&nbsp;&nbsp;<%= Replace(strBilling_cardnumber, "X", "") %>
				
			</td>
			<td>
				<%= strBilling_address %> 
				<% if strBilling_address2 <> "" then %>
					<%= strBilling_address2 %>
				<% end if %>
				&nbsp;&nbsp;&nbsp;&nbsp;
				<%= strBilling_city %>,&nbsp;<%= Replace(strBilling_state,"|","") %>&nbsp;<%= strBilling_zip %>
			</td>
			<td>
				<%= rsGetBillingAddresses.Fields.Item("date_added").Value %>
			</td>
			<td>
				<%= rsGetBillingAddresses.Fields.Item("last_updated").Value %>
			</td>
		</tr>		
		<%		
		rsGetBillingAddresses.MoveNext()
		Wend
		%>
		</table>
		<%
		end if 
		%>
		
		
<% If Not rsGetShippingAddresses.EOF Then %>
	<table class="table table-sm table-hover table-striped mt-4">
		<thead class="thead-dark">
			<tr>
			<th class="h5">Shipping addresses</th>
			<th>Added</th>
			<th>Last Updated</th>
			</tr>
		</thead>
<%
			While NOT rsGetShippingAddresses.EOF
			

			' Connect to Authorize.net CIM to get shipping address book information
			strGetAddress = "<?xml version=""1.0"" encoding=""utf-8""?>" _
			& "<getCustomerShippingAddressRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
			& MerchantAuthentication() _
			& "  <customerProfileId>" & rsGetCustomer.Fields.Item("cim_custid").Value & "</customerProfileId>" _
			& "  <customerAddressId>" & rsGetShippingAddresses.Fields.Item("cim_shippingid").Value & "</customerAddressId>" _
			& "</getCustomerShippingAddressRequest>"

			Set objResponseGetAddress = SendApiRequest(strGetAddress)

			' If connection is a success to address book than retrieve values and assign to variables
			If IsApiResponseSuccess(objResponseGetAddress) Then
			strShipping_first = objResponseGetAddress.selectSingleNode("/*/api:address/api:firstName").Text
			strShipping_last = objResponseGetAddress.selectSingleNode("/*/api:address/api:lastName").Text
			strShipping_company = objResponseGetAddress.selectSingleNode("/*/api:address/api:company").Text


			'Split out state from authorize.net and break it out to address 1 and address 2 fields
			split_address = Split(objResponseGetAddress.selectSingleNode("/*/api:address/api:address").Text, "|")
			strShipping_address = split_address(0)
			strShipping_address2 = split_address(1)

			strShipping_city = objResponseGetAddress.selectSingleNode("/*/api:address/api:city").Text
			strShipping_state = objResponseGetAddress.selectSingleNode("/*/api:address/api:state").Text
			strShipping_zip = objResponseGetAddress.selectSingleNode("/*/api:address/api:zip").Text
			strShipping_country = objResponseGetAddress.selectSingleNode("/*/api:address/api:country").Text
			strShipping_ID = objResponseGetAddress.selectSingleNode("/*/api:address/api:customerAddressId").Text
			End if
%>
	<tr>
		<td>
			<%= strShipping_first %>&nbsp;<%= strShipping_last %>
			<% if strShipping_company <> "" then %>
				&nbsp;&nbsp;&nbsp;<%= strShipping_company %>
			<% end if %>
			&nbsp;&nbsp;&nbsp;&nbsp;
			<%= strShipping_address %> &nbsp;&nbsp;
			<% if strShipping_address2 <> "" then %>
				<%= strShipping_address2 %>&nbsp;&nbsp;&nbsp;&nbsp;
			<% end if %>
			<%= strShipping_city %>, <%= Replace(strShipping_state,"|","") %>&nbsp;&nbsp;<%= strShipping_zip %>&nbsp;&nbsp;&nbsp;&nbsp;
			<%= strShipping_country %>
		</td>
			<td>
				<%= rsGetShippingAddresses.Fields.Item("date_added").Value %>
			</td>
			<td>
				<%= rsGetShippingAddresses.Fields.Item("last_updated").Value %>
			</td>
	</tr>
						
			<% 
			rsGetShippingAddresses.MoveNext()
			Wend %>
	</table>
	<%
			End If ' end Not rsGetShippingAddresses.EOF 
			%>
			
			


	<div class="card mt-4 w-25">
		<h5 class="card-header">Move orders to this account</h5>
		<div class="card-body">
		  <form name="form1" method="post" action="customer_SearchOrderHistory.asp">
			<input class="form-control form-control-sm" name="email" type="text" id="email" value="<%=(rsGetCustomer.Fields.Item("email").Value)%>" size="30">
			<input name="custID" type="hidden" id="custID" value="<%=(rsGetCustomer.Fields.Item("customer_ID").Value)%>">

			<button class="btn btn-sm btn-secondary mt-2" type="submit" name="Submit2">Add orders to account</button>
			</form>
		</div>
	  </div>



</div>

</body>

<!--#include file="includes/inc_scripts.asp"-->
<script type="text/javascript">
	//url to to do auto updating
	var auto_url = "customers/ajax-update-customer-profile.asp"
	
	// Write amount to charge to all charge buttons after user types it in
	$('#charge-cim').change(function(){
		$('.charge-amount').html($('#charge-cim').val());
		$('.charge-button').show();
	});
	
	// Click a charge button
	// Backorder processing
	$('.charge-button').click(function(){
		var amount = $('#charge-cim').val();
		var invoice = $('#main-id').val();
		var card_number = $('#card_number').val();
		var cim_account_id = $(this).attr("data-cim_account_id");
		var cim_billing_id = $(this).attr("data-cim_billing_id");
		
		$(this).hide();
		$('#loading-' + cim_billing_id).html('<i class="fa fa-spinner fa-2x fa-spin"></i>').show();
		
		$.ajax({
		method: "POST",
		dataType: "json",
		url: "customers/ajax-charge-cim.asp",
		data: {amount: amount, cim_account_id: cim_account_id, cim_billing_id: cim_billing_id, invoice: invoice, card_number: card_number}
		})
		.done(function( json, msg ) {
			$('#loading-' + cim_billing_id).html(json.status + '<br/>').delay(10000).fadeOut(1000);
			$('.charge-button').delay(10000).fadeIn(1000);
		})
		.fail(function(msg) {
			$('#loading-' + cim_billing_id).html('<div class="notice-red">CHARGE MAY HAVE GONE THROUGH... PLEASE CHECK AUTH.NET - ERROR PROCESSING</div><br/>').delay(10000).fadeOut(1000);
			$('.charge-button').delay(10000).fadeIn(1000);
		});
	}); // End Backorder processing
	
</script>
<script type="text/javascript" src="scripts/generic_auto_update_fields.js"></script>
<script type="text/javascript">
	auto_update(); // run function to update fields when tabbing out of them
</script>


</html>
<%
rsGetCustomer.Close()
Set rsGetCustomer = Nothing
%>
