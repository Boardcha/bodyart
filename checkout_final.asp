<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/functions/function-decode-to-utf8.asp" -->
<%
	page_title = "Bodyartforms order confirmation"
	page_description = "Bodyartforms order confirmation"
	page_keywords = ""
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="checkout/inc_order_details_google_analytics.asp"-->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<script type="text/javascript">
	document.getElementById('addon-alert').style.display = 'none';
</script>
<!--#include virtual="cart/inc_cart_main.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
<!--#include virtual="cart/inc_cart_grandtotal.asp"-->
<% 
if session("cc_status") = "approved" then ' APPROVED credit card ==========================  
%>
<!--#include virtual="checkout/inc_google_scripts.asp"--> 
<% if request.querystring("feedback-submitted") = ""  then %>
<div class="alert alert-success mb-4"> 
	<h5>Order Confirmation - Invoice # <%= Session("invoiceid") %></h5>
	
	<strong>Your order has been approved. Thank you so much for shopping with us!</strong>
	<br/><br/>
	If you have any questions or concerns about your order or any other matter, please feel free to contact us <a class="alert-link" href="/contact.asp">via e-mail</a> or by phone at (877) 223-5005.
	<br/><br/>
	<div class="mb-4">
	<a href="/index.asp" class="btn btn-outline-secondary">Back to Bodyartforms home page</a>
</div>
	<a class="btn btn-lg btn-light border border-primary" href="https://g.page/r/CVDk_0MEUfIlEAQ/review" target="_blank" >
		<img src="/images/homepage/google-icon.png" class="mr-2" style="height: 50px" /> Review us on Google
	</a>
</div>

<!--
<div class="card">
	<div class="card-header">
		<h6>Do you have a second to give us feedback about our website?</h6>
	</div>
		<div class="card-body">
			<form id="form-feedback" action="?feedback-submitted=yes" method="post">
				<div class="d-block mb-1">What device are you on?</div>
				
				<div class="custom-control custom-radio custom-control-inline">
					<input value="iPhone" type="radio" id="iPhone" name="platform" class="custom-control-input">
					<label class="custom-control-label" for="iPhone">iPhone</label>
				</div>
				<div class="custom-control custom-radio custom-control-inline">
					<input value="Android" type="radio" id="Android" name="platform" class="custom-control-input">
					<label class="custom-control-label" for="Android">Android</label>
				</div>
				<div class="custom-control custom-radio custom-control-inline">
					<input value="Tablet" type="radio" id="Tablet" name="platform" class="custom-control-input">
					<label class="custom-control-label" for="Tablet">Tablet</label>
				</div>
				<div class="custom-control custom-radio custom-control-inline">
					<input value="PC" type="radio" id="PC" name="platform" class="custom-control-input">
					<label class="custom-control-label" for="PC">PC</label>
				</div>
				<div class="custom-control custom-radio custom-control-inline">
					<input value="Mac" type="radio" id="Mac" name="platform" class="custom-control-input">
					<label class="custom-control-label" for="Mac">Mac</label>
				</div>
				<div class="custom-control custom-radio custom-control-inline">
					<input value="Chromebook" type="radio" id="Chromebook" name="platform" class="custom-control-input">
					<label class="custom-control-label" for="Chromebook">Chromebook</label>
				</div>
				<div class="form-group mt-3">
					<label for="feedback">Any comments or feedback?</label>
					<textarea class="form-control" name="feedback" id="feedback" rows="4"></textarea>
				</div>
				<button class="btn btn-purple" type="submit">Send feedback</button>
				<input type="hidden" name="email" value="<%= session("email") %>">
				<input type="hidden" name="name" value="<%= session("shipping_first") %>">
			</form>	
		</div>
</div>
-->

<% else 
%>
	<div class="card">
		<div class="card-header">
			<h5>Thanks for giving us feedback!</h5>
		</div>
		<div class="card-body">
			<div class="mb-4">
				<a href="/index.asp" class="btn btn-outline-secondary">Back to Bodyartforms home page</a>
			</div>
				<a class="btn btn-lg btn-light border border-primary" href="https://g.page/r/CVDk_0MEUfIlEAQ/review" target="_blank" >
					<img src="/images/homepage/google-icon.png" class="mr-2" style="height: 50px" /> Review us on Google
				</a>
		</div>
	</div>

<% end if %>

<!--#include virtual="checkout/inc_remove_items_from_cart.asp" -->
<!--#include virtual="checkout/inc_remove_all_sessions_cookies.asp"--> 
<%
 end if ' APPROVED credit card ======================================================== %>
<% if session("cc_status") = "declined" then  ' DECLINED credit card =========================  %>
	
		<div class="alert alert-danger">
			<h5>Order declined</h5>
		
		Unfortunately, your credit card has been declined due to this reason:
		<span class="d-block font-weight-bold"><%= session("cc_decline_reason") %></span>
		<div class="my-2">
		You can use the button below to go back and edit your information and try again, or please feel free to contact us <a class="alert-link" href="/contact.asp">via e-mail</a> or by phone at (877) 223-5005.
		</div>
		<a href="javascript: history.go(-1)" class="btn btn-outline-secondary">Go back to checkout</a>
		</div>
<% end if  ' DECLINED credit card ======================================================= %>

<% if session("cc_status") = "cash" then ' CASH ORDER ==========================  
%>
	
	<div class="alert alert-success"> 
		<h5>Order Confirmation - Invoice # <%= Session("invoiceid") %></h5>
		
		<strong>Thank you for your order! Instructions on how to send cash or money orders have been sent to you via e-mail.</strong>
		<br/><br/>
		If you have any questions or concerns about your order or any other matter, please feel free to contact us <a class="alert-link" href="/contact.asp">via e-mail</a> or by phone at (877) 223-5005.
		<br/><br/>
		<a href="/index.asp" class="btn btn-outline-secondary">Back to Bodyartforms home page</a>
	</div>
<!--#include virtual="checkout/inc_remove_items_from_cart.asp" -->
<!--#include virtual="checkout/inc_remove_all_sessions_cookies.asp"--> 
<% end if ' END CASH ORDER ======================================================== %>

<% if request.querystring("js") = "fail" and session("cc_status") <> "declined" and session("cc_status") <> "approved"  then  ' Script failure =======  %>
	
	<div class="alert alert-danger">
			<h5>Website Error</h5>
		
			Unfortunately our website is having trouble processing your order.
		<div class="my-2">
		You can use the button below to go back and edit your information and try again, or please feel free to contact us <a class="alert-link" href="/contact.asp">via e-mail</a> or by phone at (877) 223-5005.
		</div>
		<a href="javascript: history.go(-1)" class="btn btn-outline-secondary">Go back to checkout</a>
		</div>
<% end if  ' Script failure ================================================ %>

<% 

' Running main again below to save any session values in case the user hits back and then submits again
%>
<!--#include virtual="cart/inc_cart_main.asp"-->
</div>

<div style="height: 200px"></div>
<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript">
	// set cart count to 0
	$('#cart_count_text').html("");
	$('#mobile-cart-count').hide();
</script>