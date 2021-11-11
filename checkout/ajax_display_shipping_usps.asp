<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="cart/inc_cart_main.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<% 
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM tbl_shipping_notice"
Set rsGetNotice = objCmd.Execute()

' if getting value from country change javascript function
if request("country") <> "" then
	session("shipping-country") = request("country")
end if


' Show a $0 option for carts with just a gift certificate
if var_other_items = 1 then


		' Only display if country session has been set to display options
		'response.write "weight: " & session("weight") & "<br>"
		'response.write "country: " & session("shipping-country") & "<br>"

		if session("shipping-country") <> "" then

		if rsGetNotice.Fields.Item("country").Value = session("shipping-country")  then %>
		<div class="alert alert-warning w-100">
				<%= rsGetNotice.Fields.Item("shipping_notice").Value %>
		</div>
		<%
		End If ' weight notice

		if session("shipping-country") <> "USA"  then 
		%>
		<!--
		<div class="alert alert-danger w-100">
			<div class="font-weight-bold ">CORONAVIRUS (COVID-19) SHIPMENT DELAYS</div>
			Customers will most likely see longer delays in arrival times to certain areas of the world due to the spread of Coronavirus. As laws are changing day to day in each country, please check with your local government on shipments through borders. Many packages are being held longer at customs and borders during this time.
		</div>-->
		<% else %>
		<!--
		<div class="alert alert-danger w-100">
			<div class="font-weight-bold ">CORONAVIRUS (COVID-19) SHIPMENT DELAYS</div>
			Customers will possibly see longer delays in arrival times in certain areas (changes weekly) due to the spread of Coronavirus.
		</div>-->
		<%
		end if 

		DiscountSubtotal = 25

		' INTERNATIONAL -----------------------------
		if session("shipping-country") <> "USA" AND session("shipping-country") <> "Canada" then
			var_country = "International"
			sql_price = "ShippingAmount AS price"
			
			if CCur(var_subtotal_after_discounts) >= CCur(DiscountSubtotal) then
				sql_price = "(ShippingAmount - ShippingDiscount) AS price"
			end if
			
		end if

		' CANADA -----------------------------
		if session("shipping-country") = "Canada" then
			var_country = "Canada"
			sql_price = "ShippingAmount AS price" 

			if CCur(var_subtotal_after_discounts) >= CCur(DiscountSubtotal) then
				sql_price = "(ShippingAmount - ShippingDiscount) AS price"
			end if

		end if

		' USA --------------------------------
		if session("shipping-country") = "USA" then
			var_country = "USA"
			sql_price = "ShippingAmount AS price"
			
			if CCur(var_subtotal_after_discounts) >= CCur(DiscountSubtotal) then
				sql_price = "(ShippingAmount - ShippingDiscount) AS price"
			end if
		end if

		'EXP_estimated_delivery(expedited) and MAX_estimated_delivery(expedited max) variables are set in dhl-delivery-estimate.inc
		If request("country") = "USA" AND request("address")<>"" AND request("city")<>"" AND request("state")<>"" AND request("zip")<>"" Then%>
			<!--#include virtual="/dhl/dhl-delivery-estimate.inc"-->		
		<%
			EXP_estimated_delivery = getEstimatedDeliveryDate("EXP", request("address"), request("city"), request("state"), request("zip"), "")
			MAX_estimated_delivery = getEstimatedDeliveryDate("MAX", request("address"), request("city"), request("state"), request("zip"), "")
		End If
		
		sql_force_to_usps = ""
		'========== ZIP CODES THAT DHL DOES NOT DELIVER TO AND NEED TO BE FORCED TO USPS ========
		if request("zip") <> "" AND session("shipping-country") = "USA" then
			
			check_zip = request("zip")
			if _
				instr(check_zip, "96799") > 0 _
				OR instr(check_zip, "96910") > 0 _
				OR instr(check_zip, "96912") > 0 _
				OR instr(check_zip, "96913") > 0 _
				OR instr(check_zip, "96915") > 0 _
				OR instr(check_zip, "96916") > 0 _
				OR instr(check_zip, "96917") > 0 _
				OR instr(check_zip, "96919") > 0 _
				OR instr(check_zip, "96921") > 0 _
				OR instr(check_zip, "96923") > 0 _
				OR instr(check_zip, "96928") > 0 _
				OR instr(check_zip, "96929") > 0 _
				OR instr(check_zip, "96931") > 0 _
				OR instr(check_zip, "96932") > 0 _
				OR instr(check_zip, "96939") > 0 _
				OR instr(check_zip, "96940") > 0 _
				OR instr(check_zip, "96941") > 0 _
				OR instr(check_zip, "96942") > 0 _
				OR instr(check_zip, "96943") > 0 _
				OR instr(check_zip, "96944") > 0 _
				OR instr(check_zip, "96950") > 0 _
				OR instr(check_zip, "96951") > 0 _
				OR instr(check_zip, "96952") > 0 _
				OR instr(check_zip, "96960") > 0 _
				OR instr(check_zip, "96970") > 0 _
			then
				sql_force_to_usps = " AND ShippingName LIKE '%USPS%' "
			end if
		end if '==== if zip code found

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT IDShipping, " & sql_price & ", ShippingName, ShippingDesc_Public, ShippingWeight, ShippingType, ShippingDiscount, DiscountSubtotal, est_days_min, est_days_max, Shipping_ActualPrice, Shipping_NameWebDisplay, FRMSelected, EstimatedShipDays FROM dbo.TBL_ShippingMethods WHERE (" & session("weight") &" >= ShippingWeightMIN AND " & session("weight") &" <= ShippingWeight) AND ShippingType = ? " & sql_force_to_usps & " ORDER BY sortorder asc, price ASC"
		objCmd.Parameters.Append(objCmd.CreateParameter("Country",200,1,30,var_country))
		Set rsGetShippingOptions = objCmd.Execute()


		radio_selected = 0
		ship_loop = 0

		While NOT rsGetShippingOptions.EOF 
			if rsGetShippingOptions.Fields.Item("ShippingName").Value <> "ONLY gift certificate" then

		ship_loop = ship_loop + 1
		if ship_loop = 1 then
			var_shippingMethod_checkmark = "<i class=""ml-2 fa fa-lg fa-check""></i>"
		else
			var_shippingMethod_checkmark = ""
		end if 
		%>
		<label class="col-12 col-xs-12 col-sm-6 col-md-4 col-lg-6 col-xl-4 col-break1600-4 col-break1900-3 btn btn-light d-block btn-sm rounded-0 text-left shipping-method <%= var_active %>" style="border: .75em solid #fff" data-type="shipping-method">
			<div class="btn-sm btn-outline-secondary border border-secondary text-center d-block my-1">Select this method<span class="btn-selected"><%= var_shippingMethod_checkmark %></span></div>	
			<input name="shipping-option" type="radio" data-id="<%= ship_loop %>" value="<%= rsGetShippingOptions.Fields.Item("IDShipping").Value %>,<%= FormatNumber(rsGetShippingOptions.Fields.Item("price").Value,2) %>,USPS" data-price="<%= FormatNumber(rsGetShippingOptions.Fields.Item("price").Value,2) %>" <% if radio_selected = 0 then %> checked <% end if %>>
				<div class="d-block">
					<div class="d-block font-weight-bold">
						<%= FormatCurrency(rsGetShippingOptions.Fields.Item("price").Value,2) %> - 
						<%= rsGetShippingOptions.Fields.Item("ShippingName").Value %>
					</div>
					<%
					If EXP_estimated_delivery <> "" AND rsGetShippingOptions("ShippingName") = "DHL Basic mail" Then 
						estimated_delivery_output = "Estimated delivery date:<br>" & WeekDayName(WeekDay(EXP_estimated_delivery)) & ", " & MonthName(Month(EXP_estimated_delivery)) & " " & Day(EXP_estimated_delivery)
					ElseIf MAX_estimated_delivery <> "" AND rsGetShippingOptions("ShippingName") = "DHL Expedited Max" Then 
						estimated_delivery_output = "Estimated delivery date:<br>" & WeekDayName(WeekDay(MAX_estimated_delivery)) & ", " & MonthName(Month(MAX_estimated_delivery)) & " " & Day(MAX_estimated_delivery)
					Else
						estimated_delivery_output = rsGetShippingOptions("ShippingDesc_Public") 
					End If
					%>
					<%=estimated_delivery_output%>
				</div>
		</label>     
		<% 
		radio_selected = 1
			end if ' if <> "ONLY gift certificate"
		rsGetShippingOptions.MoveNext()

		Wend

		rsGetShippingOptions.ReQuery()


		if NOT rsGetShippingOptions.EOF then
		if session("shipping-state") = "TX" AND session("shipping-country") = "USA" then %>
		<!--
			<label class="col-12 col-xs-12 col-sm-6 col-md-4 col-lg-6 col-xl-4 col-break1600-4 col-break1900-3 btn btn-light d-block btn-sm rounded-0 text-left shipping-method <%= var_active %>" style="border: .75em solid #fff" data-type="shipping-method">
					<div class="btn-sm btn-outline-secondary border border-secondary text-center d-block my-1">Select this method<span class="btn-selected"><%= var_shippingMethod_checkmark %></span></div>
					<input name="shipping-option" type="radio" data-id="30" value="office" data-price="0" >
					<div class="d-block">
							<div class="d-block font-weight-bold">
									Free - In person pick up
							</div>
							Pick your order up at our Georgetown, TX office
						</div>
			</label>
			-->
		<% end if ' office pick up TX shipping only
		end if ' if NOT rsGetShippingOptions.EOF 

end if ' Only show if country session has been set
set rsGetShippingOptions = nothing

end if 

' only if gift cert is in cart
if var_giftcert = "yes" and var_other_items = 0 then
%>
	<label class="col-12 col-xs-12 col-sm-6 col-md-4 col-lg-6 col-xl-4 col-break1600-4 col-break1900-3 btn btn-light d-block btn-sm rounded-0 text-left <%= var_active %>" style="border: .75em solid #fff">
			<div class="btn-sm btn-outline-secondary border border-secondary text-center d-block my-1">Select this method</div>
			<input name="shipping-option" type="radio" value="Gift Certificate" data-price="0" >
			<div class="d-block">
					<div class="d-block font-weight-bold">
							NO SHIPPING REQUIRED
					</div>
					Digital gift certificate will be e-mailed to your recipient
				</div>
	</label>
<%
end if ' only gift cert in cart




DataConn.Close()
Set DataConn = Nothing
%>