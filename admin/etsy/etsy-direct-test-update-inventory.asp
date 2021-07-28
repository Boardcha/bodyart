<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/chilkat.asp" -->
<!--#include virtual="/admin/etsy/etsy-constants.asp" -->
<%
'=============== LINK TO CHILKAT ETSY UPDATE INVENTORY EXAMPLE CODE 
'=============== https://www.example-code.com/ASP/etsy_update_inventory_listing.asp
'=============== ETSY updateInventory() API   https://www.etsy.com/developers/documentation/reference/listinginventory

set rest = Server.CreateObject("Chilkat_9_5_0.Rest")
set oauth1 = Server.CreateObject("Chilkat_9_5_0.OAuth1")

oauth1.ConsumerKey = etsy_consumer_key
oauth1.ConsumerSecret = etsy_consumer_secret
oauth1.Token = etsy_oauth_permanent_token
oauth1.TokenSecret = etsy_oauth_permanent_token_secret
oauth1.SignatureMethod = "HMAC-SHA1"
success = oauth1.GenNonce(16)

autoReconnect = 1
tls = 1
success = rest.Connect("openapi.etsy.com",443,tls,autoReconnect)
If (success = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If

' Tell the REST object to use the OAuth1 object.
success = rest.SetAuthOAuth1(oauth1,1) 

jsonText = "[{""product_id"":5151872734,""property_values"":[{""property_id"": 100,""values"": [""14g (pair)""]}],""offerings"":[{""offering_id"":5012012149,""price"":""7.95"",""quantity"":1}]}]"

success = rest.AddQueryParam("products", JsonText)
success = rest.AddQueryParam("quantity_on_property", 100)
success = rest.AddQueryParam("price_on_property", 100)
success = rest.AddHeader("Content-Type","application/x-www-form-urlencoded")

jsonItemsResponseText = rest.FullRequestFormUrlEncoded("PUT","/v2/listings/871975957/inventory")
response.write jsonItemsResponseText
If (rest.LastMethodSuccess = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If


Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"

set jsonItemResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
success = jsonItemResponse.Load(jsonItemsResponseText)
jsonItemResponse.EmitCompact = 0

Response.Write "<pre>" & Server.HTMLEncode( jsonItemResponse.Emit()) & "</pre>"
Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"
Response.Write "<pre>" & Server.HTMLEncode( "Response status text: " & rest.ResponseStatusText) & "</pre>"
%>