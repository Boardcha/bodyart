<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="cart/inc_cart_main.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<% 

country_code = request("country_code")
zip_code = request("zip_code")
address= request("address")
city = request("city")
state = request("state")

If var_other_items = 1 Then

	If country_code <> "" Then

		DiscountSubtotal = 50

		' INTERNATIONAL -----------------------------
		If country_code <> "US" AND country_code <> "CA" Then
			var_country = "International"
		End If

		' CANADA -----------------------------
		If country_code = "CA" Then
			var_country = "Canada"
		End If

		' USA --------------------------------
		If country_code = "US" Then
			var_country = "USA"
		End If

		sql_price = "ShippingAmount AS price"	
		If CCur(var_subtotal_after_discounts) >= CCur(DiscountSubtotal) Then
			sql_price = "(ShippingAmount - ShippingDiscount) AS price"
		End If
			
		sql_force_to_usps = ""
		'========== ZIP CODES THAT DHL DOES NOT DELIVER TO AND NEED TO BE FORCED TO USPS ========
		If zip_code <> "" AND country_code = "US" Then
						
			If _
				instr(zip_code, "96799") > 0 _
				OR instr(zip_code, "96910") > 0 _
				OR instr(zip_code, "96912") > 0 _
				OR instr(zip_code, "96913") > 0 _
				OR instr(zip_code, "96915") > 0 _
				OR instr(zip_code, "96916") > 0 _
				OR instr(zip_code, "96917") > 0 _
				OR instr(zip_code, "96919") > 0 _
				OR instr(zip_code, "96921") > 0 _
				OR instr(zip_code, "96923") > 0 _
				OR instr(zip_code, "96928") > 0 _
				OR instr(zip_code, "96929") > 0 _
				OR instr(zip_code, "96931") > 0 _
				OR instr(zip_code, "96932") > 0 _
				OR instr(zip_code, "96939") > 0 _
				OR instr(zip_code, "96940") > 0 _
				OR instr(zip_code, "96941") > 0 _
				OR instr(zip_code, "96942") > 0 _
				OR instr(zip_code, "96943") > 0 _
				OR instr(zip_code, "96944") > 0 _
				OR instr(zip_code, "96950") > 0 _
				OR instr(zip_code, "96951") > 0 _
				OR instr(zip_code, "96952") > 0 _
				OR instr(zip_code, "96960") > 0 _
				OR instr(zip_code, "96970") > 0 _
			Then
				sql_force_to_usps = " AND ShippingName LIKE '%USPS%' "
			End If
		End If

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM TBL_Countries WHERE Country_UPSCode = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("country_code",200,1,30,country_code))
		Set rsCountry = objCmd.Execute()
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT IDShipping, " & sql_price & ", ShippingName, ShippingDesc_Public, ShippingWeight, ShippingType, ShippingDiscount, DiscountSubtotal, est_days_min, est_days_max, Shipping_ActualPrice, Shipping_NameWebDisplay, FRMSelected, EstimatedShipDays FROM dbo.TBL_ShippingMethods WHERE (" & session("weight") &" >= ShippingWeightMIN AND " & session("weight") &" <= ShippingWeight) AND ShippingType = ? " & sql_force_to_usps & " ORDER BY sortorder asc, price ASC"
		objCmd.Parameters.Append(objCmd.CreateParameter("Country",200,1,30,var_country))
		Set rsGetShippingOptions = objCmd.Execute()

		'EXP_estimated_delivery(expedited) and MAX_estimated_delivery(expedited max) variables are set in dhl-delivery-estimate.inc

		If country_code = "US" AND city <> "" AND address <> "" AND state <> "" AND zip_code <> "" Then%>
			<!--#include virtual="/dhl/dhl-delivery-estimate.inc"-->		
		<%
			EXP_estimated_delivery = getEstimatedDeliveryDate("EXP", address, city, state, zip_code, "")
			MAX_estimated_delivery = getEstimatedDeliveryDate("MAX", address, city, state, zip_code, "")		
		End If
		
		While NOT rsGetShippingOptions.EOF 
			If rsGetShippingOptions("ShippingName").Value <> "ONLY gift certificate" Then
				i = i + 1
				If EXP_estimated_delivery <> "" AND rsGetShippingOptions("ShippingName") = "DHL Basic mail" Then 
					estimated_delivery_output = "Estimated delivery date: " & WeekDayName(WeekDay(EXP_estimated_delivery)) & ", " & MonthName(Month(EXP_estimated_delivery)) & " " & Day(EXP_estimated_delivery)
				ElseIf MAX_estimated_delivery <> "" AND rsGetShippingOptions("ShippingName") = "DHL Expedited Max" Then 
					estimated_delivery_output = "Estimated delivery date: " & WeekDayName(WeekDay(MAX_estimated_delivery)) & ", " & MonthName(Month(MAX_estimated_delivery)) & " " & Day(MAX_estimated_delivery)
				Else
					estimated_delivery_output = Replace(Replace(rsGetShippingOptions("ShippingDesc_Public"), "<br>", ". "), vbCrlf, "")
				End If	
				options = options & "{""id"": """ & rsGetShippingOptions("IDShipping") & """, ""label"": ""$" & rsGetShippingOptions("price") & ": " & rsGetShippingOptions("ShippingName") & """, ""description"": """ & estimated_delivery_output & """},"
				'Example: {"id": "0", "label": "Free: Paid on original order", "description": ""}
			End If ' If <> "ONLY gift certificate"
			rsGetShippingOptions.MoveNext()
		Wend
		If request.cookies("OrderAddonsActive") <> "" then
			options = "{""id"": ""0"", ""label"": ""Free: Paid on original order"", ""description"": """"}," & options
		End If
		'Remove last comma
		options = Mid(options, 1, Len(options)-1)
		
		'OVERWRITE OPTIONS IF WE DON'T SHIP TO THE COUNTRY
		If rsCountry.EOF Then options = "{""id"": ""-1"", ""label"": ""WE CAN NOT SHIP TO " & country_code & """, ""description"": """"}"
		If Not rsCountry.EOF Then 
			if rsCountry("Display") = 0 Then options = "{""id"": ""-1"", ""label"": ""WE CAN NOT SHIP TO " & country_code & """, ""description"": """"}"
		End If
		set rsGetShippingOptions = nothing
	End If ' Only show If country session has been set

Else ' If no other items
	If var_giftcert = "yes" Then
		'NO SHIPPING REQUIRED
		'Digital gift certificate will be e-mailed to your recipient
		options = "{""id"": ""99"", ""label"": ""Free: NO SHIPPING REQUIRED"", ""description"": ""Digital gift certificate will be e-mailed to your recipient""}"
	End If
End If 

Response.Write "[" & options & "]"
DataConn.Close()
Set DataConn = Nothing
%>
