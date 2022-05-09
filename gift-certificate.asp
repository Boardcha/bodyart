<%@LANGUAGE="VBSCRIPT"%>
<%
	page_title = "Buy a gift certificate"
	page_description = "Bodyartforms purchase a gift certificate for body jewelry"
	page_keywords = "body jewelry, gift certificate"
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->

<%
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT ProductDetailID, price FROM ProductDetails WHERE ProductID = 2424 and active = 1 ORDER BY price asc"
	Set rs_getDetails = objCmd.Execute()
%>

<div class="display-5 mb-4">
	Bodyartforms gift certificate (E-mail delivery)
</div>


	<div class="alert alert-secondary col-md-6 col-lg-4">
	Gift certificates will be delivered to your recipient via e-mail immediately after your order is processed online.
</div>	
		<form class="needs-validation" novalidate>
			<div class="form-group col-md-6 col-lg-4 p-0">
				<label for="giftamount">Amount <span class="text-danger font-weight-bold">*</span></label>
				<select class="form-control" name="DetailID" id="giftamount">
				<% Do While Not rs_getDetails.EOF %>
					<option value="<%= rs_getDetails.Fields.Item("ProductDetailID").Value %>"><%= FormatCurrency(rs_getDetails.Fields.Item("price").Value) %></option>
				<% rs_getDetails.MoveNext()
				Loop %>
				</select>
				<div class="invalid-feedback">
						Amount is required
				</div>
			</div>

			<div class="form-group col-md-6 col-lg-4 p-0">
				<label for="recname">Recipient's name <span class="text-danger font-weight-bold">*</span></label>
				<input class="form-control" required name="name" id="recname" type="text"/>
				<div class="invalid-feedback">
						Recipient's name is required
				</div>
			</div>	
			
			<div class="form-group col-md-6 col-lg-4 p-0">
				<label for="giftemail">Recipient's e-mail <span class="text-danger font-weight-bold">*</span></label>
				<input class="form-control" required name="email" id="giftemail" type="email"/>
				<div class="invalid-feedback">
						Recipient's e-mail is required
				</div>
			</div>
			
			<div class="form-group col-md-6 col-lg-4 p-0">
				<label for="yourname">Your name <span class="text-danger font-weight-bold">*</span></label>
				<input class="form-control" required name="your-name" id="yourname" type="text"/>
				<div class="invalid-feedback">
						Your name is required
				</div>
			</div>	
						
			<div class="form-group col-md-6 col-lg-4 p-0">
				<label for="comments">Your message (<span id="message_characters"></span>)</label>
				<textarea class="form-control" name="gift-message" id="comments" maxlength="100" rows="6"></textarea>
			</div>

			<div class="form-group col-md-6 col-lg-4 p-0">
				<button class="btn btn-purple btn-block" type="submit" formaction="cart.asp" formmethod="post">Add to order</button>
			</div>
			<input type="hidden" name="ProductID" value="2424">
		</form>

<!--#include virtual="/bootstrap-template/footer.asp" -->

<script>
	// Character count
	var text_max = 100;
	$('#message_characters').html(text_max + ' characters remaining');

	$('#comments').keyup(function() {
		var text_length = $('#comments').val().length;
		var text_remaining = text_max - text_length;

		$('#message_characters').html(text_remaining + ' characters remaining');
	});

// Example starter JavaScript for disabling form submissions if there are invalid fields
(function() {
  'use strict';
  window.addEventListener('load', function() {
    // Fetch all the forms we want to apply custom Bootstrap validation styles to
    var forms = document.getElementsByClassName('needs-validation');
    // Loop over them and prevent submission
    var validation = Array.prototype.filter.call(forms, function(form) {
      form.addEventListener('submit', function(event) {
        if (form.checkValidity() === false) {
          event.preventDefault();
          event.stopPropagation();
        }
        form.classList.add('was-validated');
      }, false);
    });
  }, false);
})();
</script>