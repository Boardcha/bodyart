<% 
If request("initiator") = "google-pay" Or request("initiator") = "apple-pay" Then
	Response.ContentType = "application/json" 
End If
%>
<!--#include virtual="/Connections/taxjar.asp"-->
<!--#include virtual="/taxjar/taxjar-nexus-values.asp"-->

<%
session("amount_to_collect") = 0
session("state_tax_collectable") = 0
session("county_tax_collectable") = 0
session("city_tax_collectable") = 0
session("special_district_tax_collectable") = 0
session("combined_tax_rate") = 0

If request("initiator") = "google-pay" Or request("initiator") = "apple-pay" Then
	session("shipping_cost") = request("shipping_cost")
	session("taxable_amount") = request("taxable_amount")
End If

if request("tax_country") = "US" OR request("tax_country") = "USA" OR request("tax_country") = "United States" then
	taxjar_to_country = "US"
end if
if request("tax_country") = "Great Britain" OR request("tax_country") = "Great Britain and Northern Ireland" OR request("tax_country") = "United Kingdom" then
	taxjar_to_country = "GB"
end if

if request("state_taxed") = "yes" then

Set HttpReq = Server.CreateObject("MSXML2.ServerXMLHTTP")
HttpReq.open "POST", taxjar_url, false
HttpReq.setRequestHeader "Content-Type", "application/json"
HttpReq.setRequestHeader "x-api-version", "2022-01-24"
HttpReq.SetRequestHeader "Authorization", "Bearer " & taxjar_authorization & ""
HttpReq.Send("{" & _
	"""to_country"":""" & taxjar_to_country & """," & _
	"""to_state"":""" & request("tax_state") & """," & _
	"""to_zip"":""" & request("tax_zip") & """," & _
	"""to_street"": """ & request("tax_address") & """," & _
	"""from_country"":""US""," & _
	"""from_state"":""TX""," & _
	"""from_city"":""Georgetown""," & _
	"""from_zip"":""78626""," & _
	"""from_street"": ""1966 South Austin Avenue""," & _
	"""shipping"":""" & session("shipping_cost") & """," & _
	"""amount"":""" & session("taxable_amount") & """," & _
	"""line_items"": [{" & _
		"""id"":""1""," & _
		"""quantity"": 1," & _
		"""unit_price"": " & session("taxable_amount") & "," & _
		"""discount"": 0" & _
	"}]," & _
	"""nexus_addresses"": [" & _
		taxjar_nexus_values & _
	"]" & _
	"}")


'response.write HttpReq.responseText

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

	'response.write "<BR>TAXJAR OUTPUT amount_to_collect - " & session("amount_to_collect")
	'response.write "<BR>TAXJAR OUTPUT state_tax_collectable - " & session("state_tax_collectable")
	'response.write "<BR>TAXJAR OUTPUT county_tax_collectable - " & session("county_tax_collectable")
	'response.write "<BR>TAXJAR OUTPUT city_tax_collectable - " & session("city_tax_collectable")
	'response.write "<BR>TAXJAR OUTPUT special_district_tax_collectable - " & session("special_district_tax_collectable")
	'response.write "<BR>TAXJAR OUTPUT combined_tax_rate - " & session("combined_tax_rate")

' Overwrite tax variable if needed
if amount_to_collect <> 0 then
	var_salesTax = amount_to_collect
end if
if request("tax_state") <> "" then
	var_salestax_state = request("tax_state")
end if


end if 'request("state_taxed") = "yes"
%>
<% If request("initiator") = "google-pay" Or request("initiator") = "apple-pay" Then%>
	<%If amount_to_collect <> "" Then%>
		{"tax":"<%=amount_to_collect%>"}
	<%Else%>
		{"tax":"0"}
	<%End If%>	
<%End If%>