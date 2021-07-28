<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/Connections/chilkat.asp" -->
<!--#include virtual="/Connections/stripe.asp" -->
<%
set rest = Server.CreateObject("Chilkat_9_5_0.Rest")

bTls = 1
port = 443
bAutoReconnect = 1
success = rest.Connect(stripe_base_url,port,bTls,bAutoReconnect)
If (success <> 1) Then
    Response.Write "<pre>" & Server.HTMLEncode( "ConnectFailReason: " & rest.ConnectFailReason) & "</pre>"
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"

End If

success = rest.SetAuthBasic(STRIPE_SECRET_KEY,"")

success = rest.AddQueryParam("success_url","https://bodyartforms.com/stripe-success.asp")
success = rest.AddQueryParam("cancel_url","https://bodyartforms.com/stripe-cancel.asp")
success = rest.AddQueryParam("mode","payment")
success = rest.AddQueryParam("payment_method_types[0]","card")
success = rest.AddQueryParam("allow_promotion_codes","true")



success = rest.AddQueryParam("[shipping_rates[0]]","shr_1IpwJdCDpnhhWpMDA2Eww5qa")

success = rest.AddQueryParam("shipping_address_collection[allowed_countries[0]]","US")
success = rest.AddQueryParam("shipping_address_collection[allowed_countries[1]]","CA")

success = rest.AddQueryParam("line_items[0][price_data][currency]","usd")
success = rest.AddQueryParam("line_items[0][price_data][product_data][name]","Test item 0")
success = rest.AddQueryParam("line_items[0][price_data][product_data][images[0]]","http://bodyartforms-products.bodyartforms.com/thumbnail-5262021202156-38351.jpg")
success = rest.AddQueryParam("line_items[0][price_data][unit_amount]","2000")
success = rest.AddQueryParam("line_items[0][quantity]","2")
success = rest.AddQueryParam("line_items[0][adjustable_quantity][enabled]","true")
success = rest.AddQueryParam("line_items[0][adjustable_quantity][minimum]","1")
success = rest.AddQueryParam("line_items[0][adjustable_quantity][maximum]","20")


success = rest.AddQueryParam("line_items[0][dynamic_tax_rates[0]]","txr_1IvoQZCDpnhhWpMDIDS256Ow")
success = rest.AddQueryParam("line_items[0][dynamic_tax_rates[1]]","txr_1IvohUCDpnhhWpMD8AbjS4IT")
success = rest.AddQueryParam("line_items[0][dynamic_tax_rates[2]]","txr_1Ivoi0CDpnhhWpMDo7Rt1d7P")

'success = rest.AddQueryParam("line_items[1][price_data][currency]","usd")
'success = rest.AddQueryParam("line_items[1][price_data][product_data][name]","Test item 1")
'success = rest.AddQueryParam("line_items[1][price_data][product_data][images[0]]","http://bodyartforms-products.bodyartforms.com/thumbnail-5262021202156-38351.jpg")
'success = rest.AddQueryParam("line_items[1][price_data][unit_amount]","999")
'success = rest.AddQueryParam("line_items[1][quantity]","1")

'success = rest.AddQueryParam("line_items[2][price_data][currency]","usd")
'success = rest.AddQueryParam("line_items[2][price_data][product_data][name]","Test item 2")
'success = rest.AddQueryParam("line_items[2][price_data][product_data][images[0]]","http://bodyartforms-products.bodyartforms.com/thumbnail-5262021202156-38351.jpg")
'success = rest.AddQueryParam("line_items[2][price_data][unit_amount]","95")
'success = rest.AddQueryParam("line_items[2][quantity]","1")





strResponseBody = rest.FullRequestFormUrlEncoded("POST","/v1/checkout/sessions")

If (rest.LastMethodSuccess <> 1) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"

End If

set jsonResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
success = jsonResponse.Load(strResponseBody)

response.write strResponseBody

stripe_checkout_session_id = jsonResponse.StringOf("id")
%>
