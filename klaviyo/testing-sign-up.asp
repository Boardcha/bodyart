<%@LANGUAGE="VBSCRIPT"%>
	<%
	page_title = "Test Klaviyo sign ups"
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
		<!--#include virtual="/bootstrap-template/filters.asp" -->





					<div class="p-0 m-0 h5">
							GET NOTIFIED ABOUT SALES!
						</div>
						<div class="small mb-1">Sign up for our newsletter</div>
						<div class="input-group">
								<input type="text" class="form-control bg-lightgrey text-dark border-0 " placeholder="E-mail address" aria-label="E-mail address" type="text" name="homepage_newsletter_email" id="homepage_newsletter_email"  />
								<div class="input-group-append">
								  <button class="btn btn-info bg-info input-group-text px-1 text-white border-0 event-newsletter" type="button" id="homepage-newsletter-signup"><i class="fa fa-paper-plane mr-2"></i> Sign Up</button>
								</div>
							  </div>
							  <div class="mt-1" id="homepage-newsletter-msg"></div>



<script src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript">
    
// Homepage newsletter signup
$("#homepage-newsletter-signup").on("click", function () {
  $("#homepage-newsletter-signup").html('<i class="fa fa-spinner fa-2x fa-spin"></i>');
  $('#homepage_newsletter_email').hide();

  $.ajax({
      method: "post",
      dataType: "json",
      url: "/klaviyo/for-testing-klaviyo-subscribe-newsletter.asp?email=" + $('#homepage_newsletter_email').val()
      })
      .done(function(json) {
          if ($.isEmptyObject(json)) {
            $("#homepage-newsletter-msg").html('<span class="alert alert-success m-0 p-2">Thanks for signing up!</span>').show();
            $("#homepage-newsletter-signup").hide();
        } 
        if ($.isArray(json)) {
            if ((json[0].id) != "") {
                $("#homepage-newsletter-msg").html('<div class="alert alert-info m-0 p-2">You are already subscribed to our newsletter.</div>').show();
                $("#homepage-newsletter-signup").hide();
            }
        } else {
            if ((json.detail) != "") {
                $("#homepage-newsletter-msg").html('<div class="alert alert-danger m-0 p-2">' + json.detail + '</div>').show().delay(5000).fadeOut("slow");
                $("#homepage-newsletter-signup").html('Sign Up!');
                $('#homepage_newsletter_email').show();
            }

        }     
      })
      .fail(function(json) {			
          $("#homepage-newsletter-msg").html('<div class="alert alert-danger">Website ajax error</div>').show();
          $("#homepage-newsletter-signup").html('Sign Up!');
          $('#homepage_newsletter_email').show();
      })
});
</script>