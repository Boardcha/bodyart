<%@LANGUAGE="VBSCRIPT"%>
<%
	page_title = "Body jewelry measurement guide"
	page_description = "Body jewelry measurement guide"
	page_keywords = "body jewelry, measurements"
	
	var_sizing_type = "all"
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
	<div class="display-5 mb-4">
			How body jewelry is measured
	</div>
	

	<!--#include virtual="/misc_pages/measurement_help.asp" -->
<div class="my-2">
  <a class="btn btn-purple text-light" href="/images/sizing/gauge-measurement-card-printable.pdf">Click here for a printable card</a>
</div>
  <img class="border border-dark img-fluid" src="/images/sizing/gauge-measurement-card.png" alt="Measurement chart"/>  

	
	
	



<!--#include virtual="/bootstrap-template/footer.asp" --> 
<script type="text/javascript">
	$(document).on("click", "#toggle-above-1inch", function(){
		$('#above-1inch').toggle();
	})
</script>
<script src="/js/jquery.fancybox.min.js"></script>
