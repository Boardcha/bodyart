<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/admin/etsy-v3/etsy-refresh-token.asp" -->
<%
set rest = Server.CreateObject("Chilkat_9_5_0.Rest")

autoReconnect = 1
tls = 1
success = rest.Connect("openapi.etsy.com",443,tls,autoReconnect)
If (success = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If

set sbAuthHeaderVal = Server.CreateObject("Chilkat_9_5_0.StringBuilder")
success = sbAuthHeaderVal.Append("Bearer ")
success = sbAuthHeaderVal.Append(etsy_access_token)
rest.Authorization = sbAuthHeaderVal.GetAsString()

'==== IMPORTANT! We need to update qty/price value in relevant variation and send back whole array (all variants of the product) to Etsy.
' If you send only one variant, it deletes other variants from Etsy. This is how Etsy API works: get the inventory blob, edit it, and then re-submit it.
' And to clarify: In Etsy terminology, a listing refers to a product, a product refers to a variant on our end

'We need to get the array that includes all product variants from Etsy each time, another variant may be updated via ajax!
success = rest.AddQueryParam("client_id", etsy_consumer_key) 
jsonResponseText = rest.FullRequestNoBody("GET","/v3/application/listings/" & request.form("listingid") & "/inventory")

If (rest.LastMethodSuccess = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If 

set json = Server.CreateObject("Chilkat_9_5_0.JsonObject")
success = json.Load(jsonResponseText)

json.EmitCompact = 0

're-organize the json object that we'll submit to Etsy
'we need to update relevant varaint price/qty value
'and delete some elements from the json because it doesn't accept with the exact structure we got from Etsy.
For i = 0 To json.SizeOfArray("products") - 1
	Set subJson = json.ObjectOf("products[" & i & "]")
		subJson.Delete("product_id")
		subJson.Delete("is_deleted")
	Set subJson = json.ObjectOf("products[" & i & "].property_values[0]")
		subJson.Delete("scale_name")
	Set subJson = json.ObjectOf("products[" & i & "].offerings[0]")
		subJson.Delete("is_deleted") 
		subJson.Delete("offering_id")
		variant_price = subJson.IntOf("price.amount")
		'price comes with a division parameter from Etsy which is 100
		If subJson.IntOf("price.divisor") > 0 Then variant_price = subJson.IntOf("price.amount") / subJson.IntOf("price.divisor") 
		subJson.Delete("price")
		success = subJson.UpdateNumber("price", variant_price)
Next
arrayPath = "products"
relativePath = "sku"
value = request.form("sku")
caseSensitive = 0

Set product = json.FindRecord(arrayPath,relativePath,value,caseSensitive)
If (json.LastMethodSuccess <> 1) Then
    Response.Write "<pre>" & Server.HTMLEncode( "Record not found.") & "</pre>"
    Response.End
End If
success = product.SetIntOf("offerings[0].quantity", request.form("qty"))
success = product.SetNumberOf("offerings[0].price", request.form("price")) 

'Response.Write json.Emit()
'Response.End
success = rest.ClearAllQueryParams()
success = rest.AddHeader("Content-Type", "application/json")
success = rest.AddHeader("x-api-key", etsy_consumer_key)
success = rest.AddHeader("Accept","application/json")

jsonResponseText = rest.FullRequestString("PUT", "/v3/application/listings/" & request.form("listingid") & "/inventory", json.Emit())
If (rest.LastMethodSuccess = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
	Response.Status = "500 Internal Server Error"
    Response.End
End If

set jsonUpdateResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
success = jsonUpdateResponse.Load(jsonResponseText)
jsonUpdateResponse.EmitCompact = 0

Response.Write "<pre>" & Server.HTMLEncode( jsonUpdateResponse.Emit()) & "</pre>"
Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"

Set rest = Nothing
Set sbAuthHeaderVal = Nothing
Set json = Nothing
Set subJson = Nothing
Set product = Nothing
Set objCmd = Nothing
DataConn.Close
%>