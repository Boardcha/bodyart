<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/chilkat.asp" -->
<!--#include virtual="/admin/etsy/etsy-constants.asp" -->
<%
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

'response.write "<br>Listing id " & request.form("listingid")
'response.write "<br>Product id " & request.form("productid")
'response.write "<br>Offering id " & request.form("offeringid")
'response.write "<br>Sku " & request.form("sku")

' Tell the REST object to use the OAuth1 object.
success = rest.SetAuthOAuth1(oauth1,1) 

success = rest.AddQueryParam("products", "[{""product_id"":" & request.form("productid") & ",""property_values"":""[]"",""offerings"":[{""offering_id"":" & request.form("offeringid") & ",""quantity"":1}]}]")
success = rest.AddHeader("Content-Type","application/x-www-form-urlencoded")

jsonResponseText = rest.FullRequestNoBody("PUT","/v2/listings/" & request.form("listingid") & "/inventory")
If (rest.LastMethodSuccess = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If

set jsonUpdateResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
success = jsonUpdateResponse.Load(jsonResponseText)
jsonUpdateResponse.EmitCompact = 0

Response.Write "<pre>" & Server.HTMLEncode( jsonUpdateResponse.Emit()) & "</pre>"
Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"


%>