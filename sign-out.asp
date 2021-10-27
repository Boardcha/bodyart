<% @LANGUAGE="VBSCRIPT" %>
<%
	page_title = "Sign out"
	page_description = "Sign out of your Bodyartforms account"
	page_keywords = ""
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->

<%
session("storeCredit_used") = 0
session("usecredit") = ""

Response.Cookies("ID") = "" 
Response.Cookies("pass") = ""


' Variable for access to modify shipping & billing information in auth.net CIM system
session("custID_account") = ""

' put includes below because cookies have now been emptied out
%>
<!--#include virtual="cart/generate_guest_id.asp"-->
<div class="display-5 mb-5">
	Signed out
</div>
	<a href="index.asp">Go to Bodyartforms home page</a>
	<br />
	<br />
	<a href="#" data-toggle="modal" data-target="#signin">Log in again</a>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>


<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript">
	$(document).ready(function() {
		$('#cart_count_text').hide();
		$('.logged-in').hide();
		$('.logged-out').show();
	});
	</script>