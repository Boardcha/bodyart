<%@LANGUAGE="VBSCRIPT"%>
<%
    var_sponsor = request.querystring("sponsor")
	page_title = var_sponsor & " and Bodyartforms"
	page_description = ""
	page_keywords = ""
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->

<!--#include virtual="/bootstrap-template/filters.asp" -->


<div class="display-5">
		<%= Sanitize(var_sponsor) %>
	</div>
	Page info will change out depending on what the querystring has for the sponsor variable
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
