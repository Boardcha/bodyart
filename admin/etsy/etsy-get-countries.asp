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


' Tell the REST object to use the OAuth1 object.
success = rest.SetAuthOAuth1(oauth1,1)   

jsonCountryText = rest.FullRequestNoBody("GET","/v2/countries")
If (rest.LastMethodSuccess = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If

        
set jsonCountries = Server.CreateObject("Chilkat_9_5_0.JsonObject")
success = jsonCountries.Load(jsonCountryText)
jsonCountries.EmitCompact = 0


'Response.Write "<pre>" & Server.HTMLEncode( jsonCountries.Emit()) & "</pre>"

'Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"

i = 0
count_i = jsonCountries.SizeOfArray("results")
Do While i < count_i
    jsonCountries.I = i
    country = jsonCountries.StringOf("results[i].iso_country_code")
    response.write "<br/>" & country
    i = i + 1
Loop
%>