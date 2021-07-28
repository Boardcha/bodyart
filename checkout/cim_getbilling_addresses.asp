<%
if var_grandtotal > 0 then

' FIND OUT IF CUSTOMER IS REGISTERED AND HAVE DIFFERENT CHECKOUT ----------------------
if CustID_Cookie <> "" and CustID_Cookie <> 0 then 

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM TBL_Addressbook WHERE custID = ? AND address_type = 'billing' ORDER BY default_billing DESC"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
		Set rsGetBillingAddresses = objCmd.Execute()
%>
<div class="mt-4">
<div class="btn-group btn-group-toggle flex-wrap w-100" data-toggle="buttons">
<%
If Not rsGetBillingAddresses.EOF Or Not rsGetBillingAddresses.BOF Then
var_checked = "checked"
var_active = "active"

f = 0 
Do While NOT rsGetBillingAddresses.EOF

' Connect to Authorize.net CIM to get shipping address book information
		strGetBillingAddress = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<getCustomerPaymentProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <customerProfileId>" & session("cim_accountNumber") & "</customerProfileId>" _
		& "  <customerPaymentProfileId>" & rsGetBillingAddresses.Fields.Item("cim_shippingid").Value & "</customerPaymentProfileId>" _
		& "</getCustomerPaymentProfileRequest>"
		
		Set objResponseGetAddress = SendApiRequest(strGetBillingAddress)

		' If connection is a success to address book than retrieve values and assign to variables
		If IsApiResponseSuccess(objResponseGetAddress) Then

			strBilling_cardnumber = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:payment/api:creditCard/api:cardNumber").Text
			strBilling_exp = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:payment/api:creditCard/api:expirationDate").Text
			
			'Split out state from authorize.net and break it out to address 1 and address 2 fields
			split_address_billing = Split(objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:address").Text, "|")
    			strBilling_address = split_address_billing(0)
				strBilling_address2 = split_address_billing(1)

			strBilling_first = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:firstName").Text
			strBilling_last = objResponseGetAddress.selectSingleNode("/*/api:paymentProfile/api:billTo/api:lastName").Text
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
		End if

		if f = 0 then
		var_billing_checkmark = "<i class=""ml-2 fa fa-lg fa-check""></i>"
	else
		var_billing_checkmark = ""
	end if 
%>
<label class="col-12 col-xs-12 col-sm-6 col-md-4 col-lg-6 col-xl-4 col-break1600-4 col-break1900-3 btn btn-light d-block btn-sm rounded-0 text-left billing <%= var_active %>" id="billing-block-<%= rsGetBillingAddresses.Fields.Item("cim_shippingid").Value %>" style="border: .75em solid #fff"  id="<%= strBilling_ID %>" data-type="billing">
		<div class="btn-sm btn-outline-secondary border border-secondary text-center d-block my-1">Select this card<span class="btn-selected"><%= var_billing_checkmark %></span></div>
		<div class="d-block">
				<% if rsGetBillingAddresses.Fields.Item("nickname").Value <> "" then %>
					<%= rsGetBillingAddresses.Fields.Item("nickname").Value %><br/>
			   <% end if %>
			   Card ending in <strong><span id="cardid_<%= rsGetBillingAddresses.Fields.Item("cim_shippingid").Value %>"><%= Replace(strBilling_cardnumber, "X", "") %></span></strong>
		</div>
<input type="radio" name="cim_billing" value="<%= strBilling_ID %>" class="radio_billing" <%= var_checked %>>
<button class="edit-link-billing btn-sm btn btn-outline-info small py-0 my-1" type="button" data-id="<%= rsGetBillingAddresses.Fields.Item("cim_shippingid").Value %>" data-firstname="<%= strBilling_first %>" data-lastname="<%= strBilling_last %>" data-address="<%= strBilling_address %>" data-address2="<%= strBilling_address2 %>" data-city="<%= strBilling_city %>" data-state="<%= Replace(strBilling_state,"|","") %>" data-zip="<%= strBilling_zip %>" data-country="<%= strBilling_country %>" data-card="<%= strBilling_cardnumber %>">Edit</button>
<span class="badge badge-success ml-3" style="display:none" id="billing-msg-<%= rsGetBillingAddresses.Fields.Item("cim_shippingid").Value %>">Updated</span>
	</label>
          <% 
var_checked = ""
var_active = ""
f = f + 1
rsGetBillingAddresses.MoveNext()
Loop

Set rsGetBillingAddresses = Nothing	

	var_no_bill_addresses = "false"

else ' if recordset is empty set javascript to load up form
	var_no_bill_addresses = "true"

end if %>
<label  class="col-12 col-xs-12 col-sm-6 col-md-4 col-lg-6 col-xl-4 col-break1600-4 col-break1900-3 btn btn-light d-block btn-sm rounded-0 text-left billing <%= var_active %>" style="border: .75em solid #fff" id="cim_cash_click" data-type="billing" <%= hide_non_registered %>>
	<div class="btn-sm btn-outline-secondary border border-secondary text-center d-block my-1">Pay with money order or cash<span class="btn-selected"></span></div>
	<input type="radio" name="cim_billing" id="cim_cash" value="cash">
</label>
</div><!-- toggle button group-->
</div><!-- wrapper -->
<% end if  ' if customer is logged in 

end if ' if var_grandtotal > 0
%>