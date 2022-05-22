<%@LANGUAGE="VBSCRIPT"%>
<%
	page_title = "Newsletter signup"
	page_description = "Bodyartforms newsletter signup page"
	page_keywords = ""
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->


<div class="border-bottom pb-1 h5">
    STAY CONNECTED
</div>
<div class="py-2">Sign up for our newsletter and get notified anytime we run sales or special events</div>
<form name="ccoptin" target="_blank" class="mb-5">
    <div class="form-group mb-2">
        <input class="form-control" placeholder="E-mail address" type="text" name="footer_newsletter_email" id="footer_newsletter_email" />
    </div>
    <span class="btn btn-purple event-newsletter" id="footer-newsletter-signup">Sign Up!</span><span  id="footer-newsletter-msg"></span>
</form>



<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript">

	
</script>