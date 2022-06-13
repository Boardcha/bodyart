<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/taxjar.asp"-->
<!--#include virtual="/taxjar/taxjar-nexus-values.asp"-->
<html>
	<body>
	<% 
		Set HttpReq = Server.CreateObject("MSXML2.ServerXMLHTTP")
		HttpReq.open "POST", taxjar_url, false
		HttpReq.setRequestHeader "Content-Type", "application/json"
		HttpReq.setRequestHeader "x-api-version", "2022-01-24"
		HttpReq.SetRequestHeader "Authorization", "Bearer " & taxjar_authorization & ""
		HttpReq.Send("{" & _
			"""to_country"":""US""," & _
			"""to_state"":""TX""," & _
			"""to_zip"":""78626""," & _
			"""to_street"": ""1966 S. Austin Ave.""," & _
			"""from_country"":""US""," & _
			"""from_state"":""TX""," & _
			"""from_city"":""Georgetown""," & _
			"""from_zip"":""78626""," & _
			"""from_street"": ""1966 South Austin Avenue""," & _
			"""shipping"":""10""," & _
			"""amount"":""15""," & _
			"""line_items"": [{" & _
				"""product_tax_code"": """"," & _
				"""id"":""1""," & _
				"""quantity"": 1," & _
				"""unit_price"": 10," & _
				"""discount"": 0" & _
			"}," & _
			"{" & _
				"""product_tax_code"": """"," & _
				"""id"":""2""," & _
				"""quantity"": 1," & _
				"""unit_price"": 10," & _
				"""discount"": 0" & _
			"}]," & _				
			"""nexus_addresses"": [" & _
				taxjar_nexus_values & _
			"]" & _
			"}")
			
		response_cleaned = HttpReq.responseText
		Dim regEx
		Set regEx = New RegExp
		regEx.Global = true
		regEx.IgnoreCase = True
		regEx.Pattern = "[^A-Za-z0-9,_:.]"
		response_cleaned = regEx.Replace(response_cleaned, "")

		response_cleaned = replace(response_cleaned,"tax:", "")
		response_cleaned = replace(response_cleaned,"breakdown:", "")
		response_cleaned = replace(response_cleaned,"jurisdictions:", "")

		tax_array = Split(response_cleaned, ",")
		for each x in tax_array
			'response.write "X-" & x & ", "

				if instr(x,"amount_to_collect") > 0 then
					amount_to_collect = Split(x, ":")(1)
					session("amount_to_collect") = Split(x, ":")(1)
				end if
				if instr(x,"state_tax_collectable") > 0 then
					state_tax_collectable = Split(x, ":")(1)
					session("state_tax_collectable") = Split(x, ":")(1)
				end if
				if instr(x,"county_tax_collectable") > 0  then
					county_tax_collectable = Split(x, ":")(1)
					session("county_tax_collectable") = Split(x, ":")(1)
				end if
				if instr(x,"city_tax_collectable") > 0  then
					city_tax_collectable = Split(x, ":")(1)
					session("city_tax_collectable") = Split(x, ":")(1)
				end if
				if instr(x,"special_district_tax_collectable") > 0 then
					special_district_tax_collectable = Split(x, ":")(1)
					session("special_district_tax_collectable") = Split(x, ":")(1)
				end if
				if instr(x,"combined_tax_rate") > 0 AND instr(x,"line_items") <= 0 AND instr(x,"shipping") <= 0 then
					combined_tax_rate = Split(x, ":")(1)
					session("combined_tax_rate") = Split(x, ":")(1)
				end if
		next
		set HttpReq = Nothing

		Response.Write amount_to_collect
			
	%> 
	</body>
</html>