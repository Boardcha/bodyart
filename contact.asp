<%@LANGUAGE="VBSCRIPT"%>
<%
	page_title = "Contact us"
	page_description = "Contact Bodyartforms via phone or form"
	page_keywords = ""
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->

<!--#include virtual="/bootstrap-template/filters.asp" -->

<div class="display-5">
	Contact us
</div>


<form class="mt-4 col-auto col-md-6 p-0 m-0" name="form_contact" id="form_contact" method="post" action="#">
		<div class="form-group">
				<label for="reason">Reason for contact:</label>
	<select class="form-control" name="reason" id="reason">
		<option selected="selected">Select one ...</option>
		<option value="Order - Change/Update">Order - Change/Update</option>
		<option value="Order - Problem">Order - Problem</option>
		<option value="Order - General question">Order - General question</option>
		<option value="Order - Return">Order - Return</option>
		<option value="Order - Tracking status">Order - Tracking status</option>
		<option value="Product question">Product question</option>
		<option value="Custom quote">Custom quote</option>
		<option value="Website - Account issues">Website - Account issues</option>
		<option value="Website - Bug/Problem">Website - Bug/Problem</option>
		<option value="Website - Feedback">Website - Feedback</option>
		<option value="Bodyartforms contact">Other</option>
	</select>
	
	</div>
	<div class="form-group">
	<label for="name">Your name:</label>
	<input class="form-control" name="name" type="text" required />
	</div>	
	<div class="form-group">
	<label for="email">Your email:</label>
	<input class="form-control" name="email" type="email" required />
	</div>	
	<div class="form-group">
	<label for="invoice">Invoice # (if any):</label>
	<input class="form-control" name="invoice" type="text" />
</div>
<div class="form-group">
	<label for="comments">Questions or comments:</label>
	<textarea class="form-control" name="comments" rows="8" required></textarea>
</div>
	<input class="btn btn-purple btn-block btn-submit input-submit" type="submit" value="Submit">
	
</form>

<div class="load-message hide"></div>



  
<div class="pt-4 pb-1">
		Phone lines are open Monday - Friday,  9am to 5pm central time.</div>
	<h6><i class="fa fa-phone fa-lg"></i> &nbsp;&nbsp;Customer service: &nbsp;(877) 223-5005</h6><h6><i class="fa fa-phone fa-lg"></i> &nbsp;&nbsp;Pre-orders: &nbsp;(512) 943-8654</h6>
	
	<h6 class="pt-4">Address (For online order pick up ONLY):</h6>
1966 S. Austin Ave.<br />
Georgetown, TX  78626 <br />
<a href="http://goo.gl/maps/Q4mLD" title="Google Maps link" target="_blank">Link to Google Maps</a></p>   
<iframe src="//www.youtube.com/embed/42U6-0VHz5c"  allowfullscreen></iframe> 


<iframe scrolling="no" src="https://maps.google.com/maps?f=q&amp;source=s_q&amp;hl=en&amp;geocode=&amp;q=1966+S+Austin+Ave,+Georgetown,+TX&amp;aq=0&amp;oq=1966+s.+austin+a&amp;sll=30.656666,-97.70895&amp;sspn=0.277619,0.364952&amp;ie=UTF8&amp;hq=&amp;hnear=1966+S+Austin+Ave,+Georgetown,+Texas+78626&amp;t=m&amp;z=14&amp;ll=30.625777,-97.678495&amp;output=embed"></iframe>

<a href="https://maps.google.com/maps?f=q&amp;source=embed&amp;hl=en&amp;geocode=&amp;q=1966+S+Austin+Ave,+Georgetown,+TX&amp;aq=0&amp;oq=1966+s.+austin+a&amp;sll=30.656666,-97.70895&amp;sspn=0.277619,0.364952&amp;ie=UTF8&amp;hq=&amp;hnear=1966+S+Austin+Ave,+Georgetown,+Texas+78626&amp;t=m&amp;z=14&amp;ll=30.625777,-97.678495" target="_blank"><br/>View Larger Map</a>


<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript">
	
	
	$("#form_contact").submit(function(e) {
		$(".btn-submit").hide();
		$(".load-message").show();
	
		$.ajax({
		method: "post",
		url: "misc_pages/inc_contact.asp",
		data: $("#form_contact").serialize()
		})
		.done(function(msg) {
			$(".load-message").addClass("alert alert-success").html("Contact message had been sent. We will reply as soon as we can!<br/><br/>We are closed on weekends &amp; evenings after 5pm (central time).").show();
		})
		.fail(function(msg) {
			$(".load-message").addClass("alert alert-danger").html("Error sending form").show();
			$(".btn-submit").show();
		})
		
		e.preventDefault();
		return false;
	});
	
</script>