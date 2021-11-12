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