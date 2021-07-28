// Input cursor enabled on first field of address modal pop up
$('body').on('shown.bs.modal', '#updateAddress', function () {
    $('input:visible:enabled:first', this).focus();
})

// Hide all form validation errors and messages on modal close
$('body').on('hidden.bs.modal', function () {
    $('.modal form').removeClass('was-validated'); // remove class from all forms
    $('.modal :submit, .modal :button').prop('disabled', false); // enable all buttons
    $('.modal input').val(''); // clear all form inputs
    $('#message-create-account, #message-forgot').html('');
})

// Get cookie function in javascript (not jQuery plugin)
function getCookie(cname) {
    var name = cname + "=";
    var decodedCookie = decodeURIComponent(document.cookie);
    var ca = decodedCookie.split(';');
    for(var i = 0; i <ca.length; i++) {
        var c = ca[i];
        while (c.charAt(0) == ' ') {
            c = c.substring(1);
        }
        if (c.indexOf(name) == 0) {
            return c.substring(name.length, c.length);
        }
    }
    return "";
}

// Set cookie function in javascript (not jQuery plugin)
function setCookie(cname, cvalue, exdays) {
    var d = new Date();
    d.setTime(d.getTime() + (exdays*24*60*60*1000));
    var expires = "expires="+ d.toUTCString();
    document.cookie = cname + "=" + cvalue + ";" + expires + ";path=/";
}

// Get cart count from cookie and update cart badge
if (getCookie("cartCount") > 0) {
    $(".cart-count").html(getCookie("cartCount"));
}

// Remove all empty variables in querystring on any product search using filters
$('#form-filters').submit(function(){ $('select, input').each(function(){with($(this)) if (val()=='') remove();}) });

// Check all sub boxes for a filter category
$('#filters .cat-select').on('change',function(){
    var var_name = $(this).attr("data-name");
    $('.sub-' + var_name).prop("checked" , this.checked);
});

// Auto check more than one filter if gold brands are selected (check gold material, and gold brand)

$('#filters .two-filters').on('change',function(){
    var filter_type = $(this).attr("data-filter2");
    var filter_value = $(this).attr("data-filter2-value");
    $('input[name="' + filter_type + '"][value="' + filter_value + '"]').prop("checked" , this.checked);
});

// Alters bootstrap to hide alert window on [x] close rather than removing from DOM completely
    $(function () {
        $("[data-hide]").on("click", function () {
            $(this).closest("." + $(this).attr("data-hide")).hide();
        });
    });

// Changes hamburger icon to X 
    var h = $('.hamburger')
    h.on('click', function(){
    if (h.hasClass('fa-bars')) {
        h
        .removeClass('fa-bars')
        .addClass('fa-times');
    } else {
        h
        .removeClass('fa-times')
        .addClass('fa-bars');
    }
    });

// If mobile search icon is tapped, reset hamburger menu back to default state
$(".mobile-search-icon").on("click", function () {
            $('#mobilemenu, #accountmenu-bar').collapse('hide');
            $('.hamburger').removeClass('fa-times').addClass('fa-bars');
        });
// If mobile hamburger menu is tapped, close out filters
$(".hamburger").on("click", function () {
            $('#filters, #accountmenu-bar').collapse('hide');        
});
// If account menu is opened close out other open menus
$("#mobileaccountDropdown").on("click", function () {
    $('#filters, #mobilemenu').collapse('hide'); 
    $('.hamburger').removeClass('fa-times').addClass('fa-bars');       
});

    
// Filter builder output display for user to know what all checkboxes they have selected. This is text that pops up right above the apply filters button.
$("#filters input:checkbox, #filters input:radio").on("change", function () {
    var searchIDs = [];
    $("#filters input:checked").map(function(){
        searchIDs.push($(this).attr("data-friendly"));
    });
    $("#filter-builder-text").html(searchIDs.join(", "));        
});



// On page load reset body column size for PC to accomodate open filters
if($('#filters').is(':visible')) {
	$('#body-column').addClass('col-lg-9 col-xl-10');
} else {
	$('#body-column').removeClass('col-lg-9 col-xl-10');
}

// Change size of main body column if filters is clicked and visible or not
$("#toggle-filters-mobile, #toggle-filters-pc").click(function() {
    if($('#filters').is(':visible')) {
        $('#body-column').removeClass('col-lg-9 col-xl-10');
    } else {
        $('#body-column').addClass('col-lg-9 col-xl-10');
    }
});

// Display mini cart
$(".btn-cart-load").click(function(){
    $(".cart-mini-load").load("/cart/inc_cart_display_mini_bootstrap.asp");
});
  

    $('#frm-signin').submit(function (e) {
        $(".alert-signin").hide();

        // Fetch form to apply custom Bootstrap validation
        var form = $("#frm-signin")

        if (form[0].checkValidity() === false) {
            e.preventDefault()
            e.stopPropagation()
            console.log("invalid form elements");
        } else {

        $.ajax({
            method: "post",
            dataType: "json",
            url: "/accounts/ajax_sign_in.asp",
            data: $("#frm-signin").serialize(),
            beforeSend: function () {
                $('#btn_signin').attr("disabled", "disabled");
            }
        })
            .done(function (json, msg) {
                if (json.status === "logged-in") {
                    window.location = "account.asp";
                } else {
                    $(".signin-message").html("User name or password did not match. Please try again.");
                    $(".alert-signin").show();
                    $('#btn_signin').removeAttr("disabled");
                }
            })
            .fail(function () {
                $(".signin-message").html("No account found");
                $(".alert-signin").show();
                $('#btn_signin').removeAttr("disabled");
            })
            }
        form[0].classList.add('was-validated');
        e.preventDefault();
    });
    
	// Check to see if account already exists
	$("#regEmail").blur(function(){
		var email = $("#regEmail").val();
		
		$.ajax({
		method: "POST",
		dataType: "json",
		url: "/accounts/ajax_check_duplicate_account.asp",
		data: {email: email}
		})
		.done(function( json, msg ) {
			if (json.duplicate == "yes") {
				$('#message-create-account').html('<div class="alert alert-danger">A duplicate account has been found. Please <a class="alert-link" href="#" data-toggle="modal" data-target="#signin" data-dismiss="modal">login here</a>. If you have forgotten your password you can use <a class="alert-link" href="#" data-toggle="modal" data-target="#forgotpassword" data-dismiss="modal">this link</a> to retrieve it. Or, you can create a new account below.</div>');
				$('#btn-create-account').prop('disabled', true);
			}
			if (json.duplicate == "no") {
				$('#message-create-account').html('');
				$('#btn-create-account').prop('disabled', false);
			
			}
		})
		.fail(function(json, msg) {
			console.log("ajax failed");
		});
    });

	// Compare passwords
	$("#Regpassword").blur(function(){
		var password = $("#Regpassword").val();
        var confirmPassword = $("#password_confirmation").val();

                // Check for equality with the password inputs
                if (password != confirmPassword ) {
                    $('#message-create-account').html('<div class="alert alert-danger">Passwords do not match</div>');
                    $('#btn-create-account').prop('disabled', true);
                } else {
                    $('#message-create-account').html('');
                    $('#btn-create-account').prop('disabled', false);
                }
    });
    
	// Submit form
	$("#frm-register").submit(function(e) {
        // Fetch form to apply custom Bootstrap validation
        var form = $("#frm-register")

        if (form[0].checkValidity() === false) {
            e.preventDefault()
            e.stopPropagation()
            console.log("invalid form elements");
        } else {    
        
		$.ajax({
		method: "post",
		dataType: "json",
		url: "/accounts/ajax_create_account.asp",
		data: $("#frm-register").serialize()
		})
		.done(function(json, msg) {
			if (json.duplicate == "no") {
                $('#message-create-account').html('<div class="alert alert-success">Your account has been created. <a href="account.asp">Click here</a> to go to your new account page.</div>');
                $('#btn-create-account').prop('disabled', true);
			} else {
				$('#message-create-account').html('<div class="alert alert-danger">There is already an account registered with this e-mail, or another error occurred. If you have forgotten your password you can retrieve it <a class="alert-link" href="#" data-toggle="modal" data-target="#forgotpassword" data-dismiss="modal">at this link</a></div>');
			}
		})
		.fail(function(json, xhr,textStatus,err) {			
			$('#message-create-account').html('<div class="alert alert-danger">Error occurred with form information. Please make sure all fields are filled out correctly.</div>');
		})
    }
        form[0].classList.add('was-validated');
        e.preventDefault();
    }); // end submit form
    
// Forgot password
$("#frmForgotPass").submit(function(e) {
    // Fetch form to apply custom Bootstrap validation
    var form = $("#frmForgotPass")

    if (form[0].checkValidity() === false) {
        e.preventDefault()
        e.stopPropagation()
        console.log("invalid form elements");
    } else {  
        $('#message-forgot').load("/accounts/ajax_forgot_password.asp", {email:  $('#forgot_email').val()}, function() {   
        });
    }
    form[0].classList.add('was-validated');
    e.preventDefault();
});	

// Cancel adding items from order || remove
$("#btn-cancel-addons").on("click", function () {
    Cookies.remove('OrderAddonsActive', { path: '/' });
    window.location = "/cart.asp?addons=removed";
});

// Footer newsletter signup
$("#footer-newsletter-signup").on("click", function () {
    $("#footer-newsletter-signup").html('<i class="fa fa-spinner fa-2x fa-spin"></i>');
    $('#footer_newsletter_email').hide();

    $.ajax({
        method: "post",
        dataType: "json",
        url: "/constant-contact/cc-add-contact.asp?email=" + $('#footer_newsletter_email').val()
        })
        .done(function(json) {
            if (json.action != undefined) {
                $("#footer-newsletter-msg").html('<span class="alert alert-success m-0 p-2">Thanks for signing up!</span>').show().delay(5000).fadeOut("slow");
                $("#footer-newsletter-signup").hide();
            } 
            if (json.error_message != undefined) {
                $("#footer-newsletter-msg").html('<span class="alert alert-danger m-0 p-2">Invalid email</span>').show().delay(5000).fadeOut("slow");
                $("#footer-newsletter-signup").html('Sign Up!');
                $('#footer_newsletter_email').show();
            }
        })
        .fail(function(json) {			
            $("#footer-newsletter-msg").html('<span class="alert alert-danger">Website ajax error</span>').show();
            $("#footer-newsletter-signup").html('Sign Up!');
            $('#footer_newsletter_email').show();
        })
});


// Dark mode toggle
$("#darkmode-switch").on("click", function () {
	if ($("#darkmode-switch").is(':checked')) {
        $('head').append('<link href="/CSS/baf-dark.min.css" rel="stylesheet" id="darkmode" />');
        $('link[rel=stylesheet][id="lightmode"]').prop('disabled', true);
        Cookies.set("darkmode", "on", { expires: 20*365});
	} else {
        $('head').append('<link href="/CSS/baf.min.css" rel="stylesheet" id="lightmode" />');
        $('link[rel=stylesheet][id="darkmode"]').remove();
        Cookies.set("darkmode", "off", { expires: 20*365});
	}
  });