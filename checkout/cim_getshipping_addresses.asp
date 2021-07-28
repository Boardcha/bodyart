<%
' FIND OUT IF CUSTOMER IS REGISTERED AND HAVE DIFFERENT CHECKOUT ----------------------
if CustID_Cookie <> "" and CustID_Cookie <> 0 then 

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM TBL_Addressbook WHERE custID = ? AND address_type = 'shipping' ORDER BY default_shipping DESC"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
		Set rsShipping_Addresses = objCmd.Execute()

If Not rsShipping_Addresses.EOF Or Not rsShipping_Addresses.BOF Then
' Default checked item 
var_checked = "checked" '"data-checked=""checked"""
var_active = "active"
%>
<div class="mt-4">
<div class="btn-group btn-group-toggle flex-wrap w-100" data-toggle="buttons">
<%
f = 0
Do While NOT rsShipping_Addresses.EOF

' Connect to Authorize.net CIM to get shipping address book information
		strGetAddress = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<getCustomerShippingAddressRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
		& MerchantAuthentication() _
		& "  <customerProfileId>" & session("cim_accountNumber") & "</customerProfileId>" _
		& "  <customerAddressId>" & rsShipping_Addresses.Fields.Item("cim_shippingid").Value & "</customerAddressId>" _
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
					
			If not(objResponseGetAddress.selectSingleNode("/*/api:address/api:city") is nothing) then
				strShipping_city = objResponseGetAddress.selectSingleNode("/*/api:address/api:city").Text
			end if
			If not(objResponseGetAddress.selectSingleNode("/*/api:address/api:state") is nothing) then
				strShipping_state = objResponseGetAddress.selectSingleNode("/*/api:address/api:state").Text
			end if
			If not(objResponseGetAddress.selectSingleNode("/*/api:address/api:zip") is nothing) then
				strShipping_zip = objResponseGetAddress.selectSingleNode("/*/api:address/api:zip").Text
			end if
			strShipping_country = objResponseGetAddress.selectSingleNode("/*/api:address/api:country").Text
			strShipping_ID = objResponseGetAddress.selectSingleNode("/*/api:address/api:customerAddressId").Text
		End if
		
%>
<%
shipping_value = strShipping_country & "," & rsShipping_Addresses.Fields.Item("cim_shippingid").Value & "," &  Replace(strShipping_state,"|","") & "," & Replace(strShipping_state,"|","") & "," & strShipping_zip & "," &  strShipping_city

if session("shipping-checked") = shipping_value then
	var_checked = "checked"
end if

if f = 0 then
	var_shipping_checkmark = "<i class=""ml-2 fa fa-lg fa-check""></i>"
else
	var_shipping_checkmark = ""
end if 
%>
<label class="col-12 col-xs-12 col-sm-6 col-md-4 col-lg-6 col-xl-4 col-break1600-4 col-break1900-3 btn btn-light d-block btn-sm rounded-0 text-left shipping <%= var_active %>" style="border: .75em solid #fff"  id="<%= strShipping_ID %>" data-type="shipping">
		<div class="btn-sm btn-outline-secondary border border-secondary text-center d-block my-1">Ship to this address<span class="btn-selected"><%= var_shipping_checkmark %></span></div>
		<div class="d-block">
				<%= strShipping_first %>&nbsp;<%= strShipping_last %><br/>
				<% if strShipping_company <> "" then %>
					<%= strShipping_company %><br/>
				<% end if %>
				<%= strShipping_address %><br/>
				<% if strShipping_address2 <> "" then %>
					<%= strShipping_address2 %><br/>
				<% end if %>
				<%= strShipping_city %>,&nbsp;<%= Replace(strShipping_state,"|","") %>&nbsp;<%= strShipping_zip %><br/>
				<%= strShipping_country %>
			</div>
			
<input type="radio" name="cim_shipping" value="<%= strShipping_ID %>" data-address="<%= strShipping_address %> <%= strShipping_address2 %>" data-city="<%= strShipping_city %>" data-country="<%= strShipping_country %>" data-state="<%= Replace(strShipping_state,"|","") %>" data-zip="<%= strShipping_zip %>"  <%= var_checked %>>
<button class="edit-link btn-sm btn btn-outline-info small py-0 my-1" type="button" data-firstname="<%= strShipping_first %>" data-lastname="<%= strShipping_last %>" data-company="<%= strShipping_company %>" data-address="<%= strShipping_address %>" data-address2="<%= strShipping_address2 %>" data-city="<%= strShipping_city %>" data-state="<%= Replace(strShipping_state,"|","") %>" data-zip="<%= strShipping_zip %>" data-country="<%= strShipping_country %>">Edit</button>
</label>


          <% 
	' set session variables for intial page loading of UPS options for registered customers
	session("shipping-city") = strShipping_city
	session("shipping-zip") = strShipping_zip
	session("shipping-state") = Replace(strShipping_state,"|","")
	session("shipping-country") = strShipping_country
	
var_checked = ""
var_active = ""
f = f + 1
rsShipping_Addresses.MoveNext()
Loop
%>
</div><!-- toggle button group-->
</div><!-- wrapper -->
<%
Set rsShipping_Addresses = Nothing	

	var_no_ship_addresses = "false"

else ' if recordset is empty set javascript to load up form
	var_no_ship_addresses = "true"
end if 
%>
<%'= session("shipping-city") & " " & session("shipping-zip") & " " & session("shipping-state") & " " & session("shipping-country") %>
<% end if ' if customer is logged in %>