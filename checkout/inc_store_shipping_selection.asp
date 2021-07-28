<%
if (request.form("shipping_option") <> "" or request.form("shipping-option") <> "") and request.cookies("OrderAddonsActive") = "" then


' BEGIN storage of shipping option ======================================
if request.form("shipping_option") <> "" then ' sent via ajax
	shipping_option = request.form("shipping_option")
elseif request.form("shipping-option") <> "" then
	shipping_option = request.form("shipping-option")
else
	shipping_option = var_shipping_cost
end if

	
	if Instr(shipping_option, "USPS") > 0 then
	
		shipping_select = shipping_option
		shipping_arr = split(shipping_select, ",")
		usps_id = shipping_arr(0)
		shipping_cost = CCur(shipping_arr(1))

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM TBL_ShippingMethods WHERE IDShipping = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("shipping_option",3,1,30,usps_id))
		Set rsGetusps = objCmd.Execute()

		if NOT rsGetusps.EOF then
			shipping_option = rsGetusps.Fields.Item("ShippingName").Value
			shipping_info_for_email = rsGetusps.Fields.Item("ShippingName").Value & " " & rsGetusps.Fields.Item("ShippingDesc_Public").Value
		end if
		
	end if ' if usps shipping
	
	if Instr(shipping_option, "UPS") > 0 then
		
		shipping_select = shipping_option
		shipping_arr = split(shipping_select, ",")
		ups_shipping_type = shipping_arr(0)
		shipping_cost = Ccur(shipping_arr(1))
		shipping_option = shipping_arr(2)
		shipping_info_for_email = shipping_arr(2)
		ups_shipping_actual = shipping_arr(3)
		ups_shipping_weight = shipping_arr(4)
		
	end if ' if ups shipping
	
	if shipping_option = "office" then
		shipping_option = "OFFICE PICK UP"
		shipping_info_for_email = "OFFICE PICK UP"
		shipping_cost = 0
	end if
	
' END storage of shipping option ========================================

	var_shipping_cost = shipping_cost
	var_shipping_cost_friendly = FormatCurrency(shipping_cost, -1, -2, -2, -2)

	' BUG TESTING -------------------
	'response.write shipping_cost
	
	session("shipping_cost") = shipping_cost


	session("var_email_shipping_option") = shipping_info_for_email
	
end if ' being sent via form

%>