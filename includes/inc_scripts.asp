<script type="text/javascript" src="/js/jquery-2.2.3.min.js"></script>
<script type="text/javascript" src="/cart/js/cart_mini_display.js" async></script>
<script type="text/javascript" src="/js/js.cookie.js"></script>
<script type="text/javascript">
	history.navigationMode = 'compatible';
	$(document).ready(function(){
		$('#cart-count').html(Cookies.get("cartCount"));
	});
	
	var page_name_with_querystring = document.location.href.match(/[^\/]+$/)[0];
		
	var page_name_no_querystring =  location.pathname.substring(location.pathname.lastIndexOf("/") + 1);
	
	if (page_name_no_querystring == "products.asp" || page_name_no_querystring == "productdetails.asp") {
		localStorage.setItem("item_url_last_viewed", page_name_with_querystring);
	}
	
</script>

