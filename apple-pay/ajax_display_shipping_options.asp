<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="cart/inc_cart_main.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<% 

country_code = request("country_code")
zip_code = request("zip")

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
		objCmd.CommandText = "SELECT IDShipping, " & sql_price & ", ShippingName, ShippingDesc_Public, ShippingWeight, ShippingType, ShippingDiscount, DiscountSubtotal, est_days_min, est_days_max, Shipping_ActualPrice, Shipping_NameWebDisplay, FRMSelected, EstimatedShipDays FROM dbo.TBL_ShippingMethods WHERE (" & session("weight") &" >= ShippingWeightMIN AND " & session("weight") &" <= ShippingWeight) AND ShippingType = ? " & sql_force_to_usps & " ORDER BY sortorder asc, price ASC"
		objCmd.Parameters.Append(objCmd.CreateParameter("Country",200,1,30,var_country))
		Set rsGetShippingOptions = objCmd.Execute()


		While NOT rsGetShippingOptions.EOF 
			If rsGetShippingOptions("ShippingName").Value <> "ONLY gift certificate" Then
				i = i + 1
				options = options & "{""label"": """ & rsGetShippingOptions("ShippingName") & """, ""amount"": " & rsGetShippingOptions("price") & ", ""detail"": """ & Replace(Replace(rsGetShippingOptions("ShippingDesc_Public"), "<br>", ". "), vbCrlf, "") & """, ""identifier"": """ & rsGetShippingOptions("IDShipping") & """},"
				'Example: {label: 'Expedited Shipping', amount: getShippingCosts('domestic_exp'), detail: '1-3 days', identifier: 'domestic_exp'}
			End If ' If <> "ONLY gift certificate"
			rsGetShippingOptions.MoveNext()
		Wend
		If request.cookies("OrderAddonsActive") <> "" then
			options = "{""label"": ""Free: Paid on original order"", ""amount"": 0, ""detail"": """", ""identifier"": ""0""}," & options
		End If
		'Remove last comma
		options = Mid(options, 1, Len(options)-1)
		set rsGetShippingOptions = nothing
	End If ' Only show If country session has been set

Else ' If no other items
	If var_giftcert = "yes" Then
		'NO SHIPPING REQUIRED
		'Digital gift certificate will be e-mailed to your recipient
		options = "{""label"": ""Free: NO SHIPPING REQUIRED"", ""amount"": 0, ""detail"": ""Digital gift certificate will be e-mailed to your recipient"", ""identifier"": ""99""}
	End If
End If 

Response.Write "[" & options & "]"
DataConn.Close()
Set DataConn = Nothing
%>
