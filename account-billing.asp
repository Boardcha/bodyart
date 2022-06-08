<% @LANGUAGE="VBSCRIPT" CodePage = 65001  %>
<%
'Response.Write("<br/>Charset: " & Response.Charset)
'Response.Write("<br/>CodePage: " & Response.CodePage)

'IIS should process this page as 65001 (UTF-8), responses should be 
'treated as 28591 (ISO-8859-1).
Response.CharSet = "ISO-8859-1"
Response.CodePage = 28591
%>
<%
	page_title = "Your account - Saved cards"
	page_description = "Your Bodyartforms account. Edit your profile, view orders, and more."
	page_keywords = ""
	
' Clear temporary account if admin is viewing
if request.querystring("cleartemp") = "yes" then
	response.cookies("flag-tempid") = ""
	session("admin_tempcustid") = ""
end if
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->

<!--#include virtual="/bootstrap-template/filters.asp" -->
<!--#include virtual="/Connections/authnet.asp"-->
<%
var_flagged = ""

' Pull the customer information from a cookie
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM customers  WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
		Set rsGetUser = objCmd.Execute()
		
If Not rsGetUser.EOF Or Not rsGetUser.BOF Then ' Only run this info if a match was found

' Authorize.net create customer profile ... Code for older customers that did not get a profile ID created on the newer registration page --------------------------------
	strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
	& "<createCustomerProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
	& MerchantAuthentication() _
	& "<profile>" _
	& "  <merchantCustomerId>" & CustID_Cookie & "</merchantCustomerId>" _
	& "  <email>" & rsGetUser.Fields.Item("email").Value & "</email>" _
	& "</profile>" _
	& "</createCustomerProfileRequest>"
	Set objResponseCreateProfile = SendApiRequest(strReq)

	' If succcess in created a new CIM profile ID then add that new ID to our database in the customers table
	If IsApiResponseSuccess(objResponseCreateProfile) Then
	  strCustomerProfileId = objResponseCreateProfile.selectSingleNode("/*/api:customerProfileId").Text
		
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "UPDATE customers SET cim_custid = ? WHERE customer_ID = ?"
			objCmd.Parameters.Append(objCmd.CreateParameter("cim_custid",200,1,30,Server.HTMLEncode(strCustomerProfileId)))
			objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
			objCmd.Execute()
		
	End If

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
	Saved Credit Cards
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

	
<span id="address-type" data-type="billing"></span>
<button class="btn btn-warning d-block mb-4" id="add-address" type="button" text-danger data-status="add" data-url="ajax-cim-add-address.asp" data-header="Add a new credit card" data-buttonText="Add card" data-toggle="modal" data-target="#updateAddress"><i class="fa fa-plus"></i> Add new credit card</button>


<div class="message-window"></div>
<span id="show-new"></span>	
		<%
		if not rsGetBillingAddresses.eof then
			while NOT rsGetBillingAddresses.EOF
			
			' Connect to Authorize.net CIM to get CREDIT CARD information
			strGetBillingAddress = "<?xml version=""1.0"" encoding=""utf-8""?>" _
			& "<getCustomerPaymentProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
			& MerchantAuthentication() _
			& "  <customerProfileId>" & rsGetUser.Fields.Item("cim_custid").Value & "</customerProfileId>" _
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
						
				If not(objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:city") is nothing) then
					strBilling_city = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:city").Text
				end if
				
				If not(objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:state") is nothing) then
					strBilling_state = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:state").Text
				end if
				If not(objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:zip") is nothing) then
					strBilling_zip = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:zip").Text
				end if
				strBilling_country = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:country").Text
				strBilling_ID = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:customerPaymentProfileId").Text
			Else
			'	Response.Write "The operation failed with the following errors:<br>" & vbCrLf
			'	PrintErrors(objResponseGetAddress)
			End if
			
		%>
		<div class="card <%= rsGetBillingAddresses.Fields.Item("cim_shippingid").Value %>-block d-inline-block m-md-2 my-2 account-cards">
			<div class="card-body bg-light">
				<% if rsGetBillingAddresses.Fields.Item("default_billing").Value = 1 then %>
				  <h5 class="card-title" id="default-address">DEFAULT CARD</h5>
			  <% end if %>
			  <p class="card-text <%= strBilling_ID %>-address">
				<span class="font-weight-bold">Card ending in <%= Replace(strBilling_cardnumber, "X", "") %></span><br>
				<%= strBilling_address %><br>
				<% if strBilling_address2 <> "" then %>
				<%= strBilling_address2 %><br>
				<% end if %>
				<%= strBilling_city %>, <%= Replace(strBilling_state,"|","") %>&nbsp;<%= strBilling_zip %>							
					
				</p>
				<!-- begin card links -->
			  <span class="card-link pointer text-primary edit-<%= strBilling_ID %> btn-edit" text-danger data-id="<%= strBilling_ID %>" data-first="<%= strBilling_first %>" data-last="<%= strBilling_last %>" data-company="<%= strBilling_company %>" data-address="<%= strBilling_address %>" data-address2="<%= strBilling_address2 %>" data-city="<%= strBilling_city %>" data-state="<%= Replace(strBilling_state,"|","") %>" data-zip="<%= strBilling_zip %>" data-country="<%= strBilling_country %>" data-status="update" data-url="ajax-cim-update-address.asp" data-header="Update credit card" data-buttonText="Update card" data-toggle="modal" data-target="#updateAddress">
					Edit
				</span>
				

				  <i class="card-link pointer text-primary fa fa-trash-alt delete-address"  data-id="<%= rsGetBillingAddresses.Fields.Item("cim_shippingid").Value %>" data-address="Card ending in <%= Replace(strBilling_cardnumber, "X", "") %><br/><%= strBilling_address %><br><%= strBilling_address2 %>" data-toggle="modal" data-target="#deleteModal"></i>
		
			<% if  rsGetBillingAddresses.Fields.Item("default_billing").Value = 0 then %>
						<span class="card-link pointer text-primary make-default" data-id="<%= rsGetBillingAddresses.Fields.Item("ID").Value %>">
							Make default
						</span>
						<% end if %>
						<!-- end card links -->
			</div>
		  </div>	
		
		<%		
		rsGetBillingAddresses.MoveNext()
		Wend
		end if 
		%>
		<!-- Update address modal -->
<div class="modal fade" id="updateAddress" tabindex="-1" role="dialog" aria-labelledby="headerAddress" aria-hidden="true">
	<div class="modal-dialog" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title" id="headerAddress"></h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<form class="needs-validation" id="frm-cim" data-type="" data-status="" data-url="" novalidate>
		<div class="modal-body modal-scroll-long" style="height:500px">
		  <!--#include virtual="/accounts/inc-cim-address-form.asp" -->
		  <div class="message-address-modal"></div>
		</div>
		<div class="modal-footer">
		  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
		  <button type="submit" class="btn btn-purple" id="btn-modal-address"></button>
		</div>
	</form>
	  </div>
	</div>
  </div>
<!-- end update address modal -->

<!-- being confirm deletion modal -->
<div class="modal fade" id="deleteModal" tabindex="-1" role="dialog" aria-labelledby="headDelete" aria-hidden="true">
	<div class="modal-dialog modal-sm" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title" id="headDelete">Delete Card</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div class="modal-body">
				<div id="modal-confirm-address"></div>
				<div class="message-delete-modal"></div>
		</div>
		<div class="modal-footer">
		  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
		  <button type="button" class="btn btn-danger" id="confirm-delete" data-id="" data-type="">Delete</button>
		</div>
	  </div>
	</div>
  </div>
<!-- end confirm deletion modal -->

<%
end if   'rsGetUser.EOF
%>  

<!--#include virtual="/bootstrap-template/footer.asp" -->
<!-- Postgrid API -->
<script src="/js/postgrid-customized-api.js" data-pg-key="live_pk_csP2zaBTuekcKtmRMRSi9U"></script>		
<script src="/js-pages/account-address-validation.js"></script>		

<script type="text/javascript" src="/js-pages/cim-profile-management-913_1.js"></script>