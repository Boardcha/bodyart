	// START check for a duplicate account before changing e-mail
	$("#updateEmail").change(function() {
		var account_email = $("#updateEmail").val();	
			
		$.ajax({
		method: "post",
		dataType: "json",
		url: "accounts/ajax_check_duplicate_account.asp",
		data: {email: account_email}
		})
		.done(function( json, msg ) {
			if(json.duplicate == "yes") {
				$('.message-update-email').html('<div class="alert alert-danger">An account registered under the e-mail address ' + account_email + ' already exists. Please enter a different e-email address.</div>').show();
				
				$("#updateEmail").val('');
				$("#btn-update-email").prop('disabled', true);
				
			} else {
				$("#btn-update-email").prop('disabled', false);
			}
		})
		.fail(function(json, msg) {
			$('.message-update-email').html('<div class="alert alert-danger">Site error. Pleaes make sure you entered a valid email. If you continue to have trouble, please contact customer service for assistance.</div>').show();
			$("#btn-update-email").prop('disabled', true);
			
		});		

	});		// END check for a duplicate account before changing e-mail
	

	// START update account profile name
		$('#frm-update-profile').submit(function (e) {

		// Fetch form to apply custom Bootstrap validation
		var form = $("#frm-update-profile")

		if (form[0].checkValidity() === false) {
			e.preventDefault()
			e.stopPropagation()
			console.log("invalid form elements");
		} else {
			$('.message-update-profile').html('<i class="fa fa-spinner fa-2x fa-spin"></i>').show();
		$.ajax({
		method: "post",
		url: "accounts/ajax-update-profile.asp",
		data: $("#frm-update-profile").serialize()
		})
		.done(function(msg) {
			$('.message-update-profile').html('<div class="alert alert-success">Your name has been successfully updated.</div>').show();
			$('.message-update-profile').delay(5000).fadeOut('slow');
		})
		.fail(function(msg) {
			$('.message-update-profile').html('<div class="alert alert-danger">Site error. If you continue to have trouble, please contact customer us for assistance.</div>').show();
		})
		}
		form[0].classList.add('was-validated');
		e.preventDefault();
	});  // END update account profile name	

	
	// START update account e-mail address
	$('#frm-update-email').submit(function (e) {
			// Fetch form to apply custom Bootstrap validation
			var form = $("#frm-update-email")

			if (form[0].checkValidity() === false) {
				e.preventDefault()
				e.stopPropagation()
				console.log("invalid form elements");
			} else {
				$('.message-update-email').html('<i class="fa fa-spinner fa-2x fa-spin"></i>').show();
		$.ajax({
		method: "post",
		dataType: "json",
		url: "accounts/ajax-update-email.asp",
		data: $("#frm-update-email").serialize()
		})
		.done(function(json, msg) {
				if(json.status == "success") {
					$('.message-update-email').html('<div class="alert alert-success">Your e-mail address has been successfully updated.</div>').show();
					$('.message-update-email').delay(5000).fadeOut('slow');
				} else {
					$('.message-update-email').html('<div class="alert alert-danger">Error. Please make sure you entered a valid email. If you continue to have trouble, please contact customer service for assistance.</div>').show();
				}
		})
		.fail(function(json, msg) {
			$('.message-update-email').html('<div class="alert alert-danger">Site error. Please make sure you entered a valid email. If you continue to have trouble, please contact customer service for assistance.</div>').show();
		})
		
	} // check form validity
		form[0].classList.add('was-validated');
		e.preventDefault();
	});  // END update account e-mail address		


	// START compare current passwords for match
	$("#current_password").blur(function(){
		var form = $("#frm-update-pass")

		var current_password = $("#current_password").val();
		var customer_id = $("#customer_id").val();
		
		$.ajax({
		method: "POST",
		dataType: "json",
		url: "accounts/ajax_check_matching_password.asp",
		data: {password: current_password, custid: customer_id}
		})
		.done(function( json, msg ) {
			if (json.matches == "no") {
				form[0].classList.add('was-validated');
				$('#btn-update-pass').prop('disabled', true);
				$('.message-check-pass').html('The password you entered does not match the one on file').show();
			}
			if (json.matches == "yes") {
				$('#btn-update-pass').prop('disabled', false);
				form[0].classList.remove('was-validated');
				$('.message-check-pass').html('Current password is required').hide();
			
			}
		})
		.fail(function(json, msg) {
			console.log("ajax failed");
		});
	});	// END compare current passwords for match
	
	// START update account password
	$('#frm-update-pass').submit(function (e) {
		// Fetch form to apply custom Bootstrap validation
		var form = $("#frm-update-pass")
		var first_pass = $("#first_password").val();
		var pass_confirm = $("#profile_password_confirmation").val(); 

		if (form[0].checkValidity() === false) {
			e.preventDefault()
			e.stopPropagation()
			console.log("invalid form elements");
		} else {

		if ( first_pass === pass_confirm) {
			console.log("passwords do match");
		
		$('.message-update-pass').html('<i class="fa fa-spinner fa-2x fa-spin"></i>').show();

		$.ajax({
		method: "post",
		dataType: "json",
		url: "accounts/ajax-update-password.asp",
		data: $("#frm-update-pass").serialize()
		})
		.done(function(json, msg) {
				if(json.status == "success") {
					$('.message-update-pass').html('<div class="alert alert-success">Your password has been successfully updated.</div>').show();
					$('.message-update-pass').delay(5000).fadeOut('slow');
					$('.password-fields').val('');
				} else {
					$('.message-check-pass').html('Current password is required').show();
					$('.message-update-pass').html('<div class="alert alert-danger">Error. Please make sure you entered a valid password. If you continue to have trouble, please contact customer service for assistance.</div>').show();
				}
		})
		.fail(function(json, msg) {
			$('.message-update-pass').html('<div class="alert alert-danger">Site error. Please make sure you entered a valid password. If you continue to have trouble, please contact customer service for assistance.</div>').show();
		})

		} else { // if passwords do not match
			$('.message-update-pass').html('<div class="alert alert-danger pt-2">Your passwords do not match. Please try again.</div>').show();
			$('.message-update-pass').delay(5000).fadeOut('slow');
		}
		
	} // check form validity
		form[0].classList.add('was-validated');
		e.preventDefault();
	});  // END update account password		