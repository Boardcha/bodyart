<!--#include virtual="/Connections/klaviyo.asp" -->
<%
' make sure this information only fires off once per session
if session("google_sent") <> "yes" then
%>
<script type="text/javascript">
// Revised in 2021 to GA4 variable names
window.dataLayer = window.dataLayer || [];
window.dataLayer.push({
	'event':'purchase',
  'ecommerce': {
	'items': [
<%
 ' array loop for detail items to send to google
SumLineItem = 0
LineItem = 0
var_fb_line_item = ""
while NOT rsGoogle_GetOrderDetails.eof
%>
	{
		'item_name': '<%= Server.HTMLEncode(rsGoogle_GetOrderDetails.Fields.Item("title").Value) %>',
		'item_id': '<%= rsGoogle_GetOrderDetails.Fields.Item("Detailid").Value %>',
		'price': '<%= FormatNumber(rsGoogle_GetOrderDetails.Fields.Item("item_price").Value,2) %>',
		'item_brand': '<%= Server.HTMLEncode(rsGoogle_GetOrderDetails.Fields.Item("brandname").Value) %>',
		'item_category': '<%= trim(rsGoogle_GetOrderDetails.Fields.Item("jewelry").Value) %>',
<% if rsGoogle_GetOrderDetails.Fields.Item("variant").Value <> "" then %>
		'item_variant': '<%= trim(replace(rsGoogle_GetOrderDetails.Fields.Item("variant").Value,"'", "")) %>',
		<% end if %>
		'quantity': <%= rsGoogle_GetOrderDetails.Fields.Item("qty").Value %>
	},
<% 
	LineItem = rsGoogle_GetOrderDetails.Fields.Item("item_price").Value * rsGoogle_GetOrderDetails.Fields.Item("qty").Value
	SumLineItem = SumLineItem + LineItem
	
	
	var_fb_line_item = var_fb_line_item & "{'id': '" & rsGoogle_GetOrderDetails.Fields.Item("Detailid").Value & "', 'quantity': " & rsGoogle_GetOrderDetails.Fields.Item("qty").Value  & ", 'item_price': " & FormatNumber(LineItem,2) & "},"
	
	rsGoogle_GetOrderDetails.movenext()
	wend
	rsGoogle_GetOrderDetails.moveFirst()
	
	google_total = (SumLineItem - rsGetOrder.Fields.Item("total_preferred_discount").Value - rsGetOrder.Fields.Item("total_coupon_discount").Value - rsGetOrder.Fields.Item("total_free_credits").Value + rsGetOrder.Fields.Item("shipping_rate").Value + rsGetOrder.Fields.Item("total_sales_tax").Value - rsGetOrder.Fields.Item("total_store_credit").Value - rsGetOrder.Fields.Item("total_gift_cert").Value)

	facebook_pixel_total = (SumLineItem - rsGetOrder.Fields.Item("total_preferred_discount").Value - rsGetOrder.Fields.Item("total_coupon_discount").Value - rsGetOrder.Fields.Item("total_free_credits").Value - rsGetOrder.Fields.Item("total_store_credit").Value - rsGetOrder.Fields.Item("total_gift_cert").Value)
%>
],
	'currency': 'USD',
	'affiliation': 'Bodyartforms',
	'tax':'<%= FormatNumber(rsGetOrder.Fields.Item("total_sales_tax").Value, -1, -2, -2, -2) %>',
	'shipping': '<%= rsGetOrder.Fields.Item("shipping_rate").Value %>',
	'transaction_id': '<%= session("invoiceid") %>',
	'value': '<%= FormatNumber(google_total, -1, -2, -2, -2) %>', 
	'coupon': '<%= rsGetOrder.Fields.Item("coupon_code").Value %>'
}
});
</script>	


<script type="text/javascript">
	// Google Universal Analytics purchase tracking
	window.dataLayer = window.dataLayer || [];
	window.dataLayer.push({
	'event':'UApurchase',
	  'ecommerce': {
		'purchase': {
		  'products': [
	<%
	 ' array loop for detail items to send to google
	while NOT rsGoogle_GetOrderDetails.eof
	%>
		{
			'name': '<%= Server.HTMLEncode(rsGoogle_GetOrderDetails.Fields.Item("title").Value) %>',
			'id': '<%= rsGoogle_GetOrderDetails.Fields.Item("ProductID").Value %>',
			'price': '<%= FormatNumber(rsGoogle_GetOrderDetails.Fields.Item("item_price").Value,2) %>',
			'brand': '<%= Server.HTMLEncode(rsGoogle_GetOrderDetails.Fields.Item("brandname").Value) %>',
			'category': '<%= trim(rsGoogle_GetOrderDetails.Fields.Item("jewelry").Value) %>',
			<% if rsGoogle_GetOrderDetails.Fields.Item("variant").Value <> "" then %>
			'variant': '<%= trim(replace(rsGoogle_GetOrderDetails.Fields.Item("variant").Value,"'", "")) %>',
			<% end if %>
			'quantity': <%= rsGoogle_GetOrderDetails.Fields.Item("qty").Value %>
		},
	<% 	
		rsGoogle_GetOrderDetails.movenext()
		wend
	%>
	],
		'actionField': {
		'id': '<%= session("invoiceid") %>',
		'affiliation': 'Bodyartforms',
		'revenue': '<%= FormatNumber(google_total, -1, -2, -2, -2) %>', 
		'tax':'<%= FormatNumber(rsGetOrder.Fields.Item("total_sales_tax").Value, -1, -2, -2, -2) %>',
		'shipping': '<%= rsGetOrder.Fields.Item("shipping_rate").Value %>',
		'coupon': '<%= rsGetOrder.Fields.Item("coupon_code").Value %>'
		}
		}
	  }
	});

	// Google Universal Analytics purchase behavior analysis tracking
	window.dataLayer = window.dataLayer || [];
	window.dataLayer.push({
	'event':'UA_Checkout_Step_Purchase',
	  'ecommerce': {
		'checkout': {
			'actionField': {'step': 3}
		}
	  }
	});
</script>	


<!-- Facebook Pixel Code -->
<script>
!function(f,b,e,v,n,t,s){if(f.fbq)return;n=f.fbq=function(){n.callMethod?
n.callMethod.apply(n,arguments):n.queue.push(arguments)};if(!f._fbq)f._fbq=n;
n.push=n;n.loaded=!0;n.version='2.0';n.queue=[];t=b.createElement(e);t.async=!0;
t.src=v;s=b.getElementsByTagName(e)[0];s.parentNode.insertBefore(t,s)}(window,
document,'script','https://connect.facebook.net/en_US/fbevents.js');

fbq('init', '532347420293260');
fbq('track', 'Purchase', {
	value: '<%= FormatNumber(facebook_pixel_total, -1, -2, -2, -2) %>', currency:'USD',
	content_type: 'product',
	contents: [<%= LEFT(var_fb_line_item, (LEN(var_fb_line_item)-1)) %>]
	});
</script>
<noscript><img height="1" width="1" style="display:none"
src="https://www.facebook.com/tr?id=532347420293260&ev=PageView&noscript=1"
/></noscript>
<!-- End Facebook Pixel Code -->

<%
end if  ' session("google_sent") = ""
' make sure this information only fires off once per session
session("google_sent") = "yes"
%>

<!-- KLAVIYO ORDER PLACED PUSH BEGIN -->
<%
rsGoogle_GetOrderDetails.moveFirst
While NOT rsGoogle_GetOrderDetails.eof
	klaviyo_product_names = klaviyo_product_names & """" & Replace(rsGoogle_GetOrderDetails("title"), """", """""") & ""","
	klaviyo_categories = klaviyo_categories & """" & Trim(rsGoogle_GetOrderDetails("jewelry")) & ""","
	klaviyo_brands = klaviyo_brands & """" & rsGoogle_GetOrderDetails("brandname") & ""","

	klaviyo_items = klaviyo_items & "{" & _
			 """ProductID"": """ & rsGoogle_GetOrderDetails("ProductID") & """," & _
			 """SKU"": """ & rsGoogle_GetOrderDetails("DetailID") & """," & _
			 """ProductName"": """ & Replace(rsGoogle_GetOrderDetails("title"), """", """""") & """," & _
			 """Quantity"": " & rsGoogle_GetOrderDetails("qty") & "," & _
			 """ItemPrice"": " & rsGoogle_GetOrderDetails("item_price") & "," & _
			 """RowTotal"": " & rsGoogle_GetOrderDetails("qty") * rsGoogle_GetOrderDetails("item_price") & "," & _
			 """ProductURL"": ""https://bodyartforms.com/productdetails.asp?productid=" & rsGoogle_GetOrderDetails("ProductID") & """," & _
			 """ImageURL"": ""https://bodyartforms-products.bodyartforms.com/" & rsGoogle_GetOrderDetails("largepic") & """," & _
			 """ProductCategories"": [""" & Trim(rsGoogle_GetOrderDetails("jewelry")) & """]" & _
		   "},"
	rsGoogle_GetOrderDetails.MoveNext   
Wend

'Remove last coma from arrays
If klaviyo_product_names <> "" Then klaviyo_product_names = Mid(klaviyo_product_names, 1, LEN(klaviyo_product_names)-1)
If klaviyo_categories <> "" Then klaviyo_categories = Mid(klaviyo_categories, 1, LEN(klaviyo_categories)-1)
If klaviyo_brands <> "" Then klaviyo_brands = Mid(klaviyo_brands, 1, LEN(klaviyo_brands)-1)
If klaviyo_items <> "" Then klaviyo_items = Mid(klaviyo_items, 1, LEN(klaviyo_items)-1)
	
payload_order_placed = "{" & _
   """token"": """ & klaviyo_public_key & """," & _
   """event"": ""Placed Order""," & _
   """customer_properties"": {" & _
     """$email"": """ & rsGetOrder("email") & """," & _
     """$first_name"": """ & rsGetOrder("customer_first") & """," & _
     """$last_name"": """ & rsGetOrder("customer_last") & """," & _
     """$phone_number"": """ & rsGetOrder("phone") & """," & _
     """$address1"": """ & rsGetOrder("address") & """," & _
     """$address2"": """"," & _
     """$city"": """ & rsGetOrder("city") & """," & _
     """$zip"": """ & rsGetOrder("zip") & """," & _
     """$region"": """ & rsGetOrder("state") & """," & _
     """$country"": """ & rsGetOrder("country") & """" & _
   "}," & _
   """properties"": {" & _
     """$event_id"": """ & rsGetOrder("ID") & """," & _
     """$value"": " & var_subtotal & "," & _
     """OrderId"": """ & rsGetOrder("ID") & """," & _
     """Categories"": [" & klaviyo_categories & "]," & _
     """ItemNames"": [" & klaviyo_product_names & "]," & _
     """Brands"": [" & klaviyo_brands & "]," & _
     """DiscountCode"": """ & rsGetOrder("coupon_code") & """," & _ 
     """DiscountValue"": " & rsGetOrder("total_coupon_discount") & "," & _
     """Items"": [" & klaviyo_items & "]," & _
     """BillingAddress"": {" & _
       """FirstName"": """ & rsGetOrder("customer_first") & """," & _
       """LastName"": """ & rsGetOrder("customer_last") & """," & _
       """Company"": """"," & _
       """Address1"": """ & rsGetOrder("billing_address") & """," & _
       """Address2"": """"," & _
       """City"": """ & rsGetOrder("city") & """," & _
       """Region"": """ & rsGetOrder("state") & """," & _
       """RegionCode"": """"," & _
       """Country"": """ & rsGetOrder("country") & """," & _
       """CountryCode"": """"," & _
       """Zip"": """ & rsGetOrder("billing_zip") & """," & _
       """Phone"": """ & rsGetOrder("phone") & """" & _
     "}," & _
     """ShippingAddress"": {" & _
       """FirstName"": """ & rsGetOrder("customer_first") & """," & _
       """LastName"": """ & rsGetOrder("customer_last") & """," & _
       """Company"": """"," & _
       """Address1"": """ & rsGetOrder("address") & """," & _
       """Address2"": """"," & _
	   """City"": """ & rsGetOrder("city") & """," & _
       """Region"": """ & rsGetOrder("state") & """," & _
       """RegionCode"": """"," & _
       """Country"": """ & rsGetOrder("country") & """," & _
       """CountryCode"": """"," & _
       """Zip"": """ & rsGetOrder("zip") & """," & _
       """Phone"": """ & rsGetOrder("phone") & """" & _
     "}" & _
   "}," & _
   """time"": " & CStr(DateDiff("s", "01/01/1970 00:00:00", Now())) & _ 
 "}"


set http = Server.CreateObject("Chilkat_9_5_0.Http")
http.SetRequestHeader "Content-Type", "application/json"
http.Accept = "application/json"
Set resp = http.PostJson2("https://a.klaviyo.com/api/track", "application/json", payload_order_placed)
If (http.LastMethodSuccess = 0) Then
	'Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre><br>"
	Response.End
End If

jsonResponseStr = resp.BodyStr
'Response.Write jsonResponseStr	
'== KLAVIYO ORDER PLACED PUSH END ==

'== KLAVIYO ORDERED PRODUCT PUSH BEGIN ==
'FOR EACH LINE ITEM
rsGoogle_GetOrderDetails.MoveFirst
While NOT rsGoogle_GetOrderDetails.eof
	payload_ordered_product = "{" & _
	   """token"": """ & klaviyo_public_key & """," & _
	   """event"": ""Ordered Product""," & _
	   """customer_properties"": {" & _
		 """$email"": """ & rsGetOrder("email") & """," & _
		 """$first_name"": """ & rsGetOrder("customer_first") & """," & _
		 """$last_name"": """ & rsGetOrder("customer_last") & """" & _
	   "}," & _
	   """properties"": {" & _
		 """$event_id"": """ & rsGetOrder("ID") & """," & _
		 """$value"": " & rsGoogle_GetOrderDetails("item_price") & "," & _
		 """OrderId"": """ & rsGetOrder("ID") & """," & _
		 """ProductID"": """ & rsGoogle_GetOrderDetails("ProductID") & """," & _
		 """SKU"": """ & rsGoogle_GetOrderDetails("DetailID") & """," & _
		 """ProductName"": """ & Replace(rsGoogle_GetOrderDetails("title"), """", """""") & """," & _
		 """Quantity"": " & rsGoogle_GetOrderDetails("qty") & "," & _
		 """ProductURL"": ""https://bodyartforms.com/productdetails.asp?productid=" & rsGoogle_GetOrderDetails("ProductID") & """," & _
		 """ImageURL"": ""https://bodyartforms-products.bodyartforms.com/" & rsGoogle_GetOrderDetails("largepic") & """," & _
		 """Categories"": [""" & Trim(rsGoogle_GetOrderDetails("jewelry")) & """]," & _
		 """ProductBrand"": """ & rsGoogle_GetOrderDetails("brandname") & """" & _
	   "}," & _
	   """time"": " & CStr(DateDiff("s", "01/01/1970 00:00:00", Now())) & _ 
	 "}"
	 

	set http = Server.CreateObject("Chilkat_9_5_0.Http")
	http.SetRequestHeader "Content-Type", "application/json"
	http.Accept = "application/json"
	Set resp = http.PostJson2("https://a.klaviyo.com/api/track", "application/json", payload_ordered_product)
	If (http.LastMethodSuccess = 0) Then
		'Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
		Response.End
	End If

	jsonResponseStr = resp.BodyStr
	'Response.Write jsonResponseStr	
	
	rsGoogle_GetOrderDetails.MoveNext	
Wend  
'== KLAVIYO ORDERED PRODUCT PUSH END ==
%>