<%@LANGUAGE="VBSCRIPT"%>
<%
	page_title = "Page not found"
	page_description = "Bodyartforms page not found (404 error)"
	page_keywords = ""
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<div class="display-5 mb-4 clearfix">
		<img class="float-left" src="/images/500-error-graphic.png" width="150px">
		Sorry ...<br/>
		It's not you.<br/>
		It's us.
</div>


	You've unfortunately found a programming error and this page won't load. <br/>
	If you have a minute and want to submit this bug to us that would be an amazing help!
	<form class="pt-4" id="frm-500">
	<div class="form-group">
		<label class="h5" for="comments">Describe the page you were on:</label>
		<textarea class="form-control" name="comments" id="comments" rows="8" required></textarea>
	</div>
	<button type="submit" class="btn btn-purple">Submit</button>
	</form>
	<br/>
	<div class="load-message"></div>
	

	

<!--#include virtual="/bootstrap-template/footer.asp" -->

<script type="text/javascript">
	$("#frm-500").submit(function(e) {
	
		$.ajax({
		method: "post",
		url: "misc_pages/ajax-500-error.asp",
		data: $("#frm-500").serialize()
		})
		.done(function(msg) {
			$(".load-message").html('<div class="alert alert-success">Thanks! We will look into this bug as soon as we can.</div>').show();
		})
		
		e.preventDefault();
		return false;
	});
</script>