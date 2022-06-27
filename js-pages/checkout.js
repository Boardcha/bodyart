	// If customer is adding an item to an order that has not shipped yet, display the place order button even though shipping address and shipping type will not be selected
	if(Cookies.get("OrderAddonsActive") != '') {
		$('.checkout_button, #paypal-button-container').show();
		$('.submit_disabled').hide();
		$('.submit_disabled').html('');
	}

// On page load, initialize shipping options for CIM radio selection
	function triggerShippingSelection() {
	//	$('input[name="shipping-option"]:first').trigger("click");
		$('input[name="shipping-option"]').prop('checked', false);
		$('input[name="shipping-option"]:first').prop('checked', true);
		$('#ajax-shipping-options label:first').addClass('focus active');
		verify_shipmethod_selected();
	}
	// Trigger shipping section change, but not for CIM profile
	function triggerShippingCalc() {
		calcAllTotals();
		verify_shipmethod_selected();
	}

		
	// Set currency depending on if another country is chosen that we offer currency conversion for
function setCurrency() {

	if ($("input:radio[name='cim_shipping']").is(':checked')) {
		currency_country = $("input[name='cim_shipping']:checked").attr("data-country");
	} else {
		currency_country = $("select[name='shipping-country']").val();
	}

	if (currency_country === 'Australia') {
		currency_type = 'AUD'
	} else if (currency_country === 'Canada') {
		currency_type = 'CAD'
	} else if (currency_country === 'Japan') {
		currency_type = 'JPY'
	} else if (currency_country === 'New Zealand') {
		currency_type = 'NZD'
	} else if (currency_country === 'Denmark') {
		currency_type = 'DKK'
	} else if (currency_country === 'United Kingdom') {
		currency_type = 'GBP'
	} else if (currency_country === 'Great Britain') {
		currency_type = 'GBP'
	} else {
		currency_type = 'USD'
		currency_img = 'usa.png'
	}

		$.ajax({
		type: "post",
		url: "/template/inc-set-currency.asp",
		data: {currency: currency_type}
		})
	};

	function set_fields_by_country() {
		if ($('#shipping-country').val() === "USA") {
			$('.shipping-state, .billing-state, .hide_usa_zip').show();
			$('.shipping-province, .billing-province,  .shipping-province-canada, .billing-province-canada, .hide_inter_zip').hide();
			$('.shipping-province, .billing-province,  .shipping-province-canada, .billing-province-canada, .hide_inter_zip').val('');		
		}
		if ($('#shipping-country').val() === "Canada") {
			$('.shipping-province-canada, .billing-province-canada, .hide_inter_zip').show();
			$('.shipping-state, .billing-state, .shipping-province, .billing-province, .hide_usa_zip').hide();
			$('.shipping-state, .billing-state, .shipping-province, .billing-province, .hide_usa_zip').val('');
		}
		if ($('#shipping-country').val() != "USA" && $('#shipping-country').val() != "Canada") {
			$('.hide_inter_zip, .shipping-province, .billing-province').show();
			$('.shipping-state, .billing-state, .hide_usa_zip, .shipping-province-canada, .billing-province-canada').hide();
			$('.shipping-state, .billing-state, .hide_usa_zip, .shipping-province-canada, .billing-province-canada').val('');
		}
	}

	// CIM radio button selection changes to load shipping information
	$("input[name='cim_shipping']").on('change click', function(e){
		$('#shipping-loading').show();
		$('#ajax-shipping-options, #message-shipping').html('');

		var var_cim_shipping = $("input[name='cim_shipping']:checked").val();
		var cim_address = $(this).attr("data-address");
		var cim_state = $(this).attr("data-state");
		var cim_country = $(this).attr("data-country");
		var cim_city = $(this).attr("data-city");
		var cim_zip = $(this).attr("data-zip");

		setCurrency();


		
	//	console.log('CIM ' + var_cim_shipping + ', address ' + cim_address + ', state ' + cim_state);

		$('#load_temps').load("checkout/inc_store_state_selection.asp", {state:cim_state}, function() {});
		
		var page = $("#ajax-shipping-options");
		page.load("/checkout/ajax_display_shipping_usps.asp", {cim:var_cim_shipping, country:cim_country, zip:cim_zip, address:cim_address, city:cim_city, state:cim_state}, function(status) {
			$.get("/checkout/ajax_display_shipping_ups.asp", {cim:var_cim_shipping, country:cim_country, address: cim_address, city:cim_city, state:cim_state, zip:cim_zip}, function(html, status) {	
				page.append(html);
				triggerShippingSelection();
				calcAllTotals();
				$('#shipping-loading').hide();
			});
		});
		
	});


	// Load values into fields on page load
	$('input[name="e-mail"]').val(Cookies.get("e-mail"));
	$('input[name="shipping-first"]').val(Cookies.get("shipping-first"));
	$('input[name="shipping-last"]').val(Cookies.get("shipping-last"));
	$('input[name="shipping-company"]').val(Cookies.get("shipping-company"));
	$('input[name="shipping-address"]').val(Cookies.get("shipping-address"));
	$('input[name="shipping-address2"]').val(Cookies.get("shipping-address2"));
	$('input[name="shipping-city"]').val(Cookies.get("shipping-city"));
	$('#shipping-state').val(Cookies.get("shipping-state"));
	$('input[name="shipping-province"]').val(Cookies.get("shipping-province"));
	$('input[name="shipping-province-canada"]').val(Cookies.get("shipping-province-canada"));
	$('#shipping-country option[value="' + Cookies.get("shipping-country") + '"]').prop('selected', true);
	$('#billing-country option[value="' + Cookies.get("billing-country") + '"]').prop('selected', true);
	// clear state if not USA on page load
	if ($('#shipping-country').val() != "USA") {
		$('#shipping-state').prop('required', false);
	}
	
	$('input[name="shipping-zip"]').val(Cookies.get("shipping-zip"));
	$('input[name="shipping-phone"]').val(Cookies.get("shipping-phone"));
	$('input[name="billing-first"]').val(Cookies.get("billing-first"));
	$('input[name="billing-last"]').val(Cookies.get("billing-last"));
	$('input[name="billing-company"]').val(Cookies.get("billing-company"));
	$('input[name="billing-address"]').val(Cookies.get("billing-address"));
	$('input[name="billing-address2"]').val(Cookies.get("billing-address2"));
	$('input[name="billing-city"]').val(Cookies.get("billing-city"));
	$('#billing-state').val(Cookies.get("billing-state"));
	$('input[name="billing-province"]').val(Cookies.get("billing-province"));
	$('input[name="billing-province-canada"]').val(Cookies.get("billing-province-canada"));
	$('#billing-zip').val(Cookies.get("billing-zip"));

	// Run on page load -- needs to be below cim_shipping function above to load results
	if ($("input[name='cim_shipping']:checked").val() != '') {
		$('input[name="cim_shipping"]:first').trigger("click");
		window.onbeforeunload = function () {
			window.scrollTo(0, 0);
			}
	} 
	if($('#shipping-country').is(':visible')) {
		
		// set shipping on page load based on country drop down on form
		var varcountry = $("#shipping-country").val();
		var varzip = $("#shipping-zip").val();	
		var address = $("input[name='shipping-address']").val() + " " + $("input[name='shipping-address2']").val();
		var city = $("input[name='shipping-city']").val();
		var state = $("select[name='shipping-state']").val();
	
		setCurrency();
		var page = $("#ajax-shipping-options");
		page.load("/checkout/ajax_display_shipping_usps.asp", {country: varcountry, zip:varzip, address:address, city:city, state:state}, function(status) {
			$('#shipping-loading').hide();
			$.get("/checkout/ajax_display_shipping_ups.asp", {country: varcountry}, function(html, status) {	
				page.append(html);
				triggerShippingSelection();
				calcAllTotals();
			});
		});
	}
	set_fields_by_country();
		
	
	// Remember any field information that user types in and save to local storage
	$("input, select").change(function() {
		if ($(this).attr('name') != 'card_number' && $(this).attr('name') != 'billing-year' && $(this).attr('name') != 'billing-month') {
		Cookies.set($(this).attr('name'), $(this).val(), { expires: 30});
		
			//	localStorage.setItem($(this).attr('name'), $(this).val()); // removed because this does not work in Safari private browsing mode
		}
	});
	


	// Disable submit button if shipping radio button is not selected
	function verify_shipmethod_selected() {
	if($('#gift_cert_only').val() != 'yes') {
	if($('input[name="shipping-option"]').is(':checked')) 
		{ // $('.place_order').prop('disabled', false);
			//console.log("enabled submit");
		//	$('.place_order').show();
			$('.submit_disabled').hide();
			$('.submit_disabled').html('');
		} else 
		
		{ // $('.place_order').prop('disabled', true);
			//console.log("disabled submit");
			$('.place_order').hide();
			$('.submit_disabled').show();
			$('.submit_disabled').html('Please select a shipping method to continue checkout.<br/><br/>If you do not see any shipping options, please fill out the shipping address form entirely.');
		}
	}
	};
	
	// Display areas if user has scripts enabled
	$('.shipping-same').show();
	$('.shipping-section').show();
	$('.edit-link').show();
	$('.edit-link-billing').show();
	// auto check radio buttons
	$('input[data-checked="checked"]').prop('checked',true);
		
	
	// START if place order button is clicked
	$('#checkout_form').submit(function(e) {
						
		// Disable password required if neither field is filled out
		if($('input[type="password"]').val() === ''){
			$('.create-password').hide();
		}

		$('.checkout_button, #msg-location-replace').hide();
		$('.processing-message').show();
		$('.processing-message').html('<div class="alert alert-success mt-2"><i class="fa fa-spinner fa-2x fa-spin"></i> Please wait ... PROCESSING ORDER</div>');

		// Fetch form to apply custom Bootstrap validation
		console.log($("#shipping-country").val());
		var form = $("#checkout_form")
		if (form[0].checkValidity() === false) {
			var this_error = '';
			var all_required_errors = '';
			$('.form-control:invalid').each(function(i, obj) {
				if ($(this).attr("data-friendly-error") != undefined){
					this_error = '- ' + $(this).attr("data-friendly-error") + '<br/>';
					all_required_errors = all_required_errors + this_error;
				}
			});
			$('.processing-message').html('<div class="alert alert-danger mt-2">Some fields that are required have not been filled out. Please fix the fields that are highlighted in red.<div class="small mt-2">' + all_required_errors + '</div></div>');
			$('.checkout_button').show();
			showShippingAddressInputs();
			showBillingAddressInputs();
		} else {
		console.log("processing order...");
			$.ajax({
			method: "POST",
			dataType: "json",
			url: "checkout/ajax_process_payment.asp",
			data: $("#checkout_form").serialize()
			})
			.done(function( json, msg ) {
				
				// PAYPAL payment
				if (json.paypal === "yes" && json.store_credit != "used_partial" && json.covered_infull_giftcert != "yes") {  // PayPal payment
					console.log("Paypal ... transferring");
					$('.processing-message').html('<div class="alert alert-success mt-2">Transferring you to PayPal. Please wait ...</div>');
					$('#msg-location-replace').html('<div class="alert alert-success mt-2"><i class="fa fa-spinner fa-2x fa-spin"></i> Transferring you to PayPal. Please wait ...</div>').show();
					window.location = "/checkout-paypal-authnet-v1.asp?step=1";
				} 
				// AFTERPAY payment
				else if (json.afterpay === "yes" && json.store_credit != "used_partial" && json.covered_infull_giftcert != "yes") {  // afterpay payment
					console.log("Afterpay ... transferring");
					$('.processing-message').html('<div class="alert alert-success mt-2">Transferring you to AfterPay. Please wait ...</div>');
					$('#msg-location-replace').html('<div class="alert alert-success mt-2"><i class="fa fa-spinner fa-2x fa-spin"></i> Transferring you to AfterPay. Please wait ...</div>').show();
						$.ajax({   
							method: "post",
							dataType: "json",
							url: "afterpay/afterpay-create-checkout.asp"
							})
							.done(function(json,msg) {
								afterpay_token = json.afterpay_token;
								console.log(afterpay_token);
								AfterPay.initialize({countryCode: "US"});
								AfterPay.redirect({token: afterpay_token});
							})
							.fail(function(json,msg) {
								$('.checkout_button').show();
								$('.processing-message').html('<div class="alert alert-danger mt-2">Afterpay processing error</div>');
							})
				} 
				else if (json.cash === "yes") { // Cash payment
					console.log("Cash order");
							$('#msg-location-replace').html('<div class="alert alert-success mt-2"><i class="fa fa-spinner fa-2x fa-spin"></i> Finalizing cash order - Please wait ...</div>').show();
							window.location = "/checkout_final.asp";
				} 
				else if (json.store_credit === "used_partial") { // Store credit paid for order in full
					console.log("Store credit paid for order in full");
					$('#msg-location-replace, .processing-message').html('<div class="alert alert-success mt-2"><i class="fa fa-spinner fa-2x fa-spin"></i> Finalizing order & store credit - Please wait ...</div>').show();
					window.location = "/checkout_final.asp";
				}
				else if (json.covered_infull_giftcert === "yes") { // Gift cert paid for order in full
					console.log("Gift certificate paid for order in full");
					$('#msg-location-replace, .processing-message').html('<div class="alert alert-success mt-2"><i class="fa fa-spinner fa-2x fa-spin"></i> Finalizing order & gift certificate - Please wait ...</div>').show();
					window.location = "/checkout_final.asp";
				}
				else { // CREDIT CARD payments
					// Check to see if any stock changes have occurred
					if (json.stock_status === "fail") {
						$('.checkout_button').show();
						$('.processing-message').html('').hide();

						$('.stock-error').show();
						$('.stock-error').load("/cart/inc_stock_display_notice.asp");
						calcAllTotals();
					} else { // If items are in stock then proceed with credit card attempt
						// if credit card has been attempted to be charged
						if (json.cc_approved === "yes") {
							console.log("Payment successful");
							$('#msg-location-replace').html('<div class="alert alert-success mt-2"><i class="fa fa-spinner fa-2x fa-spin"></i> Finalizing payment - Please wait ...</div>').show();
							window.location = "/checkout_final.asp";
						} else {
							// Declined card					
							console.log("Payment declined");
							console.log("msg.responseText: " + msg.responseText);
							console.log("json.errorText: " + json.errorText);
							$('.checkout_button').show();
							$('.processing-message').html('<div class="alert alert-danger mt-2">We are sorry but the merchant has declined your card. You have not been charged. <br/><br/>The reason they gave is <span class="font-weight-bold">' + json.cc_reason + '</span></div>');
						}				
					}
				}
			})
			.fail(function() {
				$('.checkout_button').show();
				$('.processing-message').html('<div class="alert alert-danger mt-2">Unfortunately our website is having trouble processing your order.<br/><br/>Please review your information and try again or feel free to contact us <a class="alert-link" href="/contact.asp">via e-mail</a> or by phone at (877) 223-5005.<br><br><strong>ERROR MESSAGE: Ajax Failure</strong><br>' + $("#timestamp").val() + '</div>');
			});

		} // bootstrap form validation
		form[0].classList.add('was-validated');
		e.preventDefault();
		return false;
		
	}); // END if place order button is clicked
	
	// If add shipping button is clicked then show div
	$('.add-new-shipping-button').click(function() {
		// Uncheck radio button address
		$('input[name="cim_shipping"]').prop('checked', false);
		//Enable all fields
		$('.shipping-address-form input, .shipping-address-form select').attr('disabled', false);
		$('#changing-shipping-header').html("Add a new");
		$('#shipping-save-wrapper, #shipping-same-billing-wrapper').show();
		$('.shipping-address-form').show();
		$('#cancel-shipping-add').show();		
		$('.add-new-shipping-button').hide();
		$('#shipping-status').val("add");
		$('#cim_shipping_addresses, .checkout_button').hide();

		clearShippingAddressInputs();
		set_fields_by_country();
		/*	No need to fill inputs from cookies on entering a new address. Ugur.
		$('.shipping-address-form input:visible:not(#shipping-full-address), select:visible').each(function(){
				if (Cookies.get($(this).attr('name'))) {					
					$(this).val(Cookies.get($(this).attr('name')));
				}
		})*/
		$("#ajax-shipping-options").html('');
		
		if(isUserFromUSAorCANADA()){
			showShippingAddressValidation();
			clearShippingAddressInputs();
			$('#selected-shipping-address').hide();	
			$("#shipping-address-container").hide();
		}else{
			hideShippingAddressValidation();
			$("#shipping-address-container").show();
		}	
		$('#chk-shipping-manual-address-input').prop('checked', false);	
	});
	
	// If add billing button is clicked then show div
	$('.add-new-billing-button').click(function() {
		// Uncheck radio button address
		$('.radio_billing').removeAttr('checked');
		$('#cim_paypal, #cim_cash').prop('checked', false);
		$('input[name="cim_billing"]').parent('div').removeClass('notice-grey');
		//Enable all fields
		$('.billing-address-form input, .billing-address-form select').attr('disabled', false);

		if ($('#shipping-first-checkout').is(":visible")){
		} else {
			$('#shipping-same-billing-wrapper').hide();
		}
		$('#card-save-wrapper').show();
		
		// Toggle required attribute
		toggleRequiredBillingTrue();
		
		// Only show billing same as shipping checkbox if shipping form is currently open 
		if($('.shipping-address-form').is(':hidden')) {
			$('.shipping-same').hide();
		}

		$('.billing-address-form, #credit_card_inputs, .billing-input-fields').show();
		$('#card-save').show();
		$('.add-new-billing-button').hide();
		$('#cancel-billing-add').show();
		$('#billing-status').val("add");
		$('#cim_billing_addresses').hide();
		
		set_fields_by_country();	
		
		$('.billing-address-form input:visible:not(#billing-full-address), select:visible').each(function(){	
			if (Cookies.get($(this).attr('name'))) {					
					$(this).val(Cookies.get($(this).attr('name')));
				}
		})
				
		if(isUserFromUSAorCANADA()){
			if($("#billing-address").val() !="" && $("#billing-city").val() !="" && $("#billing-zip").val() !="" && $("#billing-country").val() !="" && getBillingState() != ""){
				createAddressBubble("billing", 'fromCookies');
				$("#billing-address-container").hide();
			}else{
				showBillingAddressValidation();
				clearBillingAddressInputs();
				$('#selected-billing-address').hide();	
				$("#billing-address-container").hide();
			}
		}else{
			hideBillingAddressValidation();
			$("#billing-address-container").show();
		}
		$('#chk-billing-manual-address-input').prop('checked', false);
	})
	
	// If cancel add SHIPPING button is pressed then clear entire form contents 

	$('#cancel-shipping-add').click(function() {
		// Check radio button address
		$('input[name="cim_shipping"]:radio:first').prop('checked', true);
		$('input[name="cim_shipping"]:radio:first').trigger("click");
	//	var cim_id = $("input[name='cim_shipping']:checked").val();
	//	$('#' + cim_id).addClass('active');

		//Disable all fields
		$('.shipping-address-form input, .shipping-address-form select').attr('disabled', true);
		$('.shipping-address-form').find(':input').not(':button, :submit, :reset, :hidden').val('');
		$('.shipping-address-form').hide();
		$('.add-new-shipping-button, .checkout_button, #cim_shipping_addresses').show();
		$('#cancel-shipping-add, .shipping-same').hide();
		$('#changing-shipping-header').html("Add a new");	
		$(".shipping-address-form :input").removeClass("valid error")
		$('#shipping-status').val("");
		
		$('.shipping-address-form input:hidden, select:hidden').each(function(){
			if ( localStorage[$(this).attr('name')]) {
				localStorage.removeItem($(this).attr('name'));
			}
		});
		
		// Request new shipping options when adding new address
		$('#ajax-shipping-options').load("checkout/ajax_display_shipping_usps.asp", {cim:$("input[name='cim_shipping']:checked").val()}, function() {
			triggerShippingSelection();
			calcAllTotals();
		});
	});
	
	// If cancel add BILLING button is pressed then clear entire form contents 
	$('#cancel-billing-add').click(function() {
		// Check radio button address
		$('input[name="cim_billing"]:radio:first').prop('checked', true);		
		$('input[name="cim_billing"]:radio:first').parent('div').addClass('notice-grey');
		//Disdable all fields
		$('.billing-address-form input, .billing-address-form select').attr('disabled', true);
		$('.billing-address-form input, .billing-address-form select').attr('disabled', true);
		$('.billing-address-form').find(':input').not(':button, :submit, :reset, :hidden').val('');
		$('.billing-address-form').hide();
		$('.add-new-billing-button').show();
		$('#cancel-billing-add, #btn-save-credit-card').hide();
		$(".billing-address-form :input").removeClass("valid error")
		$('#billing-status').val("");
		$('#cim_billing_addresses').show();
		
		$('.billing-address-form input:hidden, select:hidden').each(function(){
			if ( localStorage[$(this).attr('name')]) {
				localStorage.removeItem($(this).attr('name'));
			}
		});
	});	
	
	// shipping and cart total when shipping option is changed
	// .on allows manipulation of loaded page elements after DOM is loaded
	$(document).on('change', 'input[name="shipping-option"]:radio', function() {
		calcAllTotals();
		verify_shipmethod_selected();
	});

	// Save guest email if typed in
	$(document).on('change', '#e-mail', function() {
		guest_email = $('#e-mail').val();
		$.ajax({
			method: "POST",
			url: "checkout/ajax-save-guest-email.asp",
			data: {email: guest_email}
			})
	});

	
	
	
	// Load UPS results on page load, detect CIM
	var var_cim_shipping_pageLoad = $("input[name='cim_shipping']:checked").val();
	var cim_country_pageLoad = $("input[name='cim_shipping']:checked").attr("data-country");
	var cim_state_pageLoad = $("input[name='cim_shipping']:checked").attr("data-state");
	var cim_zip_pageLoad = $("input[name='cim_shipping']:checked").attr("data-zip");
	var cim_city_pageLoad = $("input[name='cim_shipping']:checked").attr("data-city");
	$('.ups-ajax').load("checkout/ajax_display_shipping_ups.asp", {cim:var_cim_shipping_pageLoad, country:cim_country_pageLoad, city:cim_city_pageLoad, state:cim_state_pageLoad, zip:cim_zip_pageLoad});
	
		
	// CIM radio button selection changes to load billing information
	$("input[name='cim_billing']").on('load change', function(){
		var var_cim_billing = $("input[name='cim_billing']:checked").val();
	});

	// Populate form if shipping edit buttons is clicked
	$('.edit-link').click(function(e) {
		$('.shipping-address-form input, .shipping-address-form select').attr('disabled', false);
		$('#shipping-status').val("update");
		$('#shipping-save-wrapper, .add-new-shipping-button, #cim_shipping_addresses').hide();
		$('#cancel-shipping-add').show();
		var edit_country = $(this).data("country");
		var edit_zip = $(this).data("zip");
		$('.shipping-address-form').show();
		$("input[name='shipping-first']").val($(this).data("firstname"));
		$("input[name='shipping-last']").val($(this).data("lastname"));
		$("input[name='shipping-address']").val($(this).data("address"));
		$("input[name='shipping-address2']").val($(this).data("address2"));
		$("input[name='shipping-company']").val($(this).data("company"));
		$("input[name='shipping-city']").val($(this).data("city"));
		$("select[name='shipping-state']").val($(this).data("state"));
		$("select[name='shipping-province-canada']").val($(this).data("state"));
		$("input[name='shipping-province']").val($(this).data("state"));
		$("input[name='shipping-zip']").val($(this).data("zip"));
		$("select[name='shipping-country']").val($(this).data("country"));	
		if (edit_country == 'USA') {		
			$('.shipping-province-canada, .shipping-province').hide();
			$('.shipping-province-canada, .shipping-province').val('');
			$('.shipping-state').show();
			$('#shipping-state').prop('required', true);
			$('#shipping-province-canada, #shipping-province').prop('required', false);
		}	
		if (edit_country == 'Canada') {
			$("select[name='shipping-province-canada']").append('<option value=' + $(this).data("state") + ' selected>' + $(this).data("state") + '</option>');
			$('.shipping-province-canada').show();
			$('.shipping-state, .shipping-province').hide();
			$('.shipping-state, .shipping-province').val('');
			$('#shipping-province-canada').prop('required', true);
			$('#shipping-state, #shipping-province').prop('required', false); 
		}	
		if (edit_country != 'USA' && edit_country != 'Canada') {
			$("input[name='shipping-province']").val($(this).data("state"));
			$('.shipping-province').show();
			$('.shipping-state, .shipping-province-canada').hide();
			$('.shipping-state, .shipping-province-canada').val('');
			$('#shipping-state, #shipping-province, #shipping-province-canada').prop('required', false);
		}	
		
		//If the address is complete, create the address bubble	
		if(isUserFromUSAorCANADA() && $("#shipping-address").val() !="" && $("#shipping-city").val() !="" && $("#shipping-zip").val() !="" && $("#shipping-country").val() !="" && getShippingState() != ""){
			createAddressBubble('shipping', 'fromCookies');		
			$("#shipping-address-container").hide();
		}else{
			showShippingAddressInputs();			
		}			
		
		$('#ajax-shipping-options').load("checkout/ajax_display_shipping_usps.asp", {country:edit_country, zip:edit_zip}, function(){
				$('input[name="shipping-option"]:first').click();
				calcAllTotals();
				document.getElementById('shipping-card').scrollIntoView(true);
		});
		e.preventDefault();
	});
	
	// If a password is typed in, check to see if an e-mail account already exists for what they typed in
	
	$("input[name='password_confirmation'], input[name='password']").change(function() {
		
		var account_email = $("#e-mail").val();
	
		$.ajax({
		method: "POST",
		dataType: "json",
		url: "accounts/ajax_check_duplicate_account.asp",
		data: {email: account_email}
		})
		.done(function( json, msg ) {
			if(json.duplicate == "yes") {
				$('#duplicate_account').show();
				$('#duplicate_account').html('An account registered under the e-mail address ' + account_email + ' already exists.');
				$("input[name='password_confirmation'], input[name='password']").val('');
				//console.log("duplicate account found");
			} else {
				$('#duplicate_account').hide();
			}
		})
		.fail(function(json, msg) {
			console.log("fail");
		});		
	});
	
	
	
		// Populate form if BILLING EDIT buttons is clicked
	$('.edit-link-billing').click(function() {
		$('.billing-input-fields').show();
		$('#credit_card_inputs').show();
		$('.billing-address-form input, .billing-address-form select').attr('disabled', false);
		$('#billing-status').val("update");
		$('#card-save-wrapper, #shipping-same-billing-container, .add-new-billing-button, #cim_billing_addresses').hide();
		$('#billing-address-autocomplete label').show();
		$('#cancel-billing-add, #btn-save-credit-card').show();
		var edit_country = $(this).data("country");
		$('.billing-address-form').show();
		$('#btn-save-credit-card').attr("data-id", $(this).data("id"));
		$("input[name='billing-first']").val($(this).data("firstname"));
		$("input[name='billing-last']").val($(this).data("lastname"));
		$("input[name='billing-address']").val($(this).data("address"));
		$("input[name='billing-address2']").val($(this).data("address2"));
		$("input[name='billing-company']").val($(this).data("company"));
		$("input[name='billing-city']").val($(this).data("city"));
		$("select[name='billing-state']").val($(this).data("state"));
		$("select[name='billing-province-canada']").val($(this).data("state"));
		$("input[name='billing-province']").val($(this).data("state"));
		$("input[name='billing-zip']").val($(this).data("zip"));
		$("select[name='billing-country']").val($(this).data("country"));

		if (edit_country == 'USA') {		
			$('#billing-province-canada, #billing-province').val('');
		}	
		if (edit_country == 'Canada') {
			$("select[name='billing-province-canada']").append('<option value=' + $(this).data("state") + ' selected>' + $(this).data("state") + '</option>');
			$('#billing-state, #billing-province').val('');	
		}	
		if (edit_country != 'USA' && edit_country != 'Canada') {
			$('#billing-state, #billing-province-canada').val('');	
			$('.billing-state, .billing-province, .billing-province-canada').prop('required', false);
		}
		//If the address is complete, create the address bubble	
		if(isUserFromUSAorCANADA() && $("#billing-address").val() !="" && $("#billing-city").val() !="" && $("#billing-zip").val() !="" && $("#billing-country").val() !="" && getBillingState() != ""){
			createAddressBubble('billing', 'fromCookies');
			$("#billing-address-container").hide();			
		}else{
			showBillingAddressInputs();
		}		
	});

	// UPDATE SAVED CREDIT CARD INFORMATION
	$('#btn-save-credit-card').click(function() {
		$('#spinner-update-billing').show();
		$('#msg-update-billing').hide();

		var cim_id = $(this).attr("data-id");
		var first = $('#billing-first').val();
		var last = $('#billing-last').val();
		var address = $('#billing-address').val();
		var address2 = $('#billing-address2').val();
		var city = $('#billing-city').val();
		var state = $('#billing-state').val();
		var province = $('#billing-province').val();
		var province_canada = $('#billing-province-canada').val();
		var zip = $('#billing-zip').val();
		var country = $('#billing-country').val();
		var card_number = $('#cardNumber').val();
		var cc_month = $('#creditCardMonth').val();
		var cc_year = $('#creditCardYear').val();

		$.ajax({
			method: "post",
			dataType: "json",
			url: "accounts/ajax-cim-update-address.asp",
			data:	{type: "billing",
					"cim-id": cim_id,
					card_number: card_number,
					"billing-month": cc_month,
					"billing-year": cc_year,
					first: first,
					last: last,
					address: address,
					address2: address2, 
					city: city,
					state: state,
					province: province,
					province_canada: province_canada,
					zip: zip,
					country: country
					}
			})
			.done(function(json, msg, data) {
				$('#spinner-update-billing').hide();
				if(json.status == "success") {
					// Check radio button address
					$('input[name="cim_billing"]:radio[value=' + cim_id + ']').prop('checked', true);
					$('input[name="cim_billing"]:radio[value=' + cim_id + ']').trigger("click");
					//Disable all fields
					$('.billing-address-form input, .billing-address-form select').attr('disabled', true);
					$('.billing-address-form input, .billing-address-form select').attr('disabled', true);
					$('.billing-address-form').find(':input').not(':button, :submit, :reset, :hidden').val('');

					document.getElementById("cardid_" +cim_id).textContent = card_number.slice(card_number.length - 4);
					// Hide and show elements
					$('.billing-address-form, #cancel-billing-add').hide();
					$('.add-new-billing-button, #cim_billing_addresses, #billing-msg-' +cim_id).show();
					$('#billing-status').val("");

					document.getElementById("billing-block-" +cim_id).scrollIntoView();
				} else { // if update failed
					$('#msg-update-billing').html(json.reason).show();
				}
			})
			.fail(function(json, msg, data) {
				$('#spinner-update-billing').hide();
				$('#msg-update-billing').html("Code error " + json.reason).show();
			});
	}); // End update saved credit card information
		
	// Billing same as shipping information checkbox
	$("input[name='shipping-same-billing']").change(function(){
		if ($("input[name='shipping-same-billing']").prop('checked')) {
			$("#billing-first").val($("#shipping-first-checkout").val());
			$("#billing-last").val($("#shipping-last-checkout").val());
			$("#billing-address").val($("#shipping-address").val());
			$("#billing-address2").val($("#shipping-address2").val());
			$("#billing-city").val($("#shipping-city").val());
			$("#billing-state").val($("#shipping-state").val());
			$("#billing-province-canada").val($("select[name='shipping-province-canada']").val());
			$("#billing-province").val($("input[name='shipping-province']").val());
			$("#billing-zip").val($("#shipping-zip").val());
			$("#billing-country").val($("select[name='shipping-country']").val());
			
			//If all shipping address fields to copy are filled, show the address bubble instead of input fields
			if($("#shipping-address").val() !="" && $("#shipping-city").val() !="" && $("#shipping-zip").val() !="" && $("#shipping-country").val() !=""){

				var state = getShippingState();
	
				if(state != "" && isUserFromUSAorCANADA()){							
					createAddressBubble('billing', 'fromShipping');
					$('#selected-billing-address').hide().fadeIn('fast');
					$("#billing-address-container").hide();
					$('#billing-country').change();		
				}else{
					showBillingAddressInputs();
				}	
			}else{
				showBillingAddressInputs();
			}			
			
			// Trigger change to save field info to local storage
			 $("#billing-first, #billing-last, #billing-address, #billing-address2, #billing-city, #billing-state, #billing-province-canada, #billing-province, #billing-zip, #billing-country").trigger('change');
			 $("#chk-billing-manual-address-input-container").hide();
		} else {
			clearBillingAddressInputs();
			$("#billing-bubble-close").trigger('click');
			if (isUserFromUSAorCANADA()){
				showBillingAddressValidation();
				$("#billing-address-container").hide();
			}else{
				showBillingAddressInputs();		
			}
		}
	}); // End shipping same as billing
	
	function isUserFromUSAorCANADA(){
		var isUserLocal = false;
		$.ajax({
		  dataType: "json",
		  url: "https://pro.ip-api.com/json/?key=1FbdfMEofSriXRs&fields=countryCode",
		  async: false, 
		  success: function(data) {
			if(data.countryCode == 'US' || data.countryCode == 'CA'){
				isUserLocal = true;
			}
		  }
		});	
		return isUserLocal;	
	}
	
	function showShippingAddressValidation(){
		$("#shipping-address-container").hide();	
		$("#shipping-address-autocomplete").show();
		$("#chk-shipping-manual-address-input-container").show();
	}
	
	function showBillingAddressValidation(){
		$('#billing-address-autocomplete').show();
		$("#billing-address-container").hide();	
		$("#chk-billing-manual-address-input-container").show();	
	}
	
	function hideShippingAddressValidation(){
		$("#shipping-address-autocomplete").hide();
		$("#chk-shipping-manual-address-input-container").hide();
	}
	
	function hideBillingAddressValidation(){
		$('#billing-address-autocomplete').hide();
		$("#chk-billing-manual-address-input-container").hide();	
	}	
	
	function showShippingAddressInputs(){
		$("#shipping-address-container").show();		
		$('#selected-shipping-address').hide();	
		$("#shipping-address-autocomplete").hide();	
		$("#chk-shipping-manual-address-input-container").hide();	
	}
	
	function showBillingAddressInputs(){
		$("#billing-address-container").show();		
		$('#selected-billing-address').hide();	
		$("#billing-address-autocomplete").hide();	
		$("#chk-billing-manual-address-input-container").hide();
	}	
	
	
	function getShippingState(){
		var state = "";
		if($("select[name='shipping-country']").val() == "USA")
			state = $("#shipping-state").val();
		else if($("select[name='shipping-country']").val() == "Canada")
			state = $("select[name='shipping-province-canada']").val();
		else
			state = $("input[name='shipping-province']").val();	
		return state;	
	}
	
	function getBillingState(){
		var state = "";
		if($("select[name='billing-country']").val() == "USA")
			state = $("#billing-state").val();
		else if($("select[name='billing-country']").val() == "Canada")
			state = $("select[name='billing-province-canada']").val();
		else
			state = $("input[name='billing-province']").val();	
		return state;	
	}	
	
	$("input[name='chk-billing-manual-address-input']").change(function(){
		if (this.checked) {
			showBillingAddressInputs();
		}
	});	
	
	$("input[name='chk-shipping-manual-address-input']").change(function(){
		if (this.checked) {
			showShippingAddressInputs();
		}
	});	
	
	function closeAddressBubble(section){
		if(section =='billing')
			closeBillingBubble();
		if(section =='shipping')
			closeShippingBubble();			
	};
	
	function closeBillingBubble(){
	  showBillingAddressValidation();
	  clearBillingAddressInputs();
	  $("#shipping-same-billing").prop('checked', false);
	  $('#chk-billing-manual-address-input').prop('checked', false);
	  $('#billing-full-address').val('');
	  $('#billing-full-address').focus();
	};	
	
	function closeShippingBubble(){
	  showShippingAddressValidation();
	  clearShippingAddressInputs();
	  $("#shipping-same-billing").prop('checked', false);
	  $('#chk-shipping-manual-address-input').prop('checked', false);
	  $('#shipping-full-address').val('');
	  $('#shipping-full-address').focus();
	};		

	function clearAddressInputs(section){
		if (section=="billing")
			clearBillingAddressInputs();
		if (section=="shipping")
			clearShippingAddressInputs();			
	}
	
	function clearBillingAddressInputs(){
		$("input[name='billing-first']").val('');
		$("input[name='billing-last']").val('');
		$("input[name='billing-address']").val('');
		$("input[name='billing-address2']").val('');
		$("input[name='billing-city']").val('');
		$("input[name='billing-state']").removeAttr("selected");
		$("input[name='billing-country']").removeAttr("selected");
		$("input[name='billing-zip']").val('');
		$("input[name='billing-province']").val('');
		$("input[name='billing-province-canada']").val('');		
		$('#selected-billing-address').hide();
	}
	
	function clearShippingAddressInputs(){
		$("#shipping-full-address").val('');
		$("input[name='shipping-address']").val('');
		$("input[name='shipping-address2']").val('');
		$("input[name='shipping-city']").val('');
		$("input[name='shipping-country']").removeAttr("selected");
		$("input[name='shipping-state']").removeAttr("selected");
		$("input[name='shipping-province-canada']").removeAttr("selected");
		$("input[name='shipping-province']").val('');
		$("input[name='shipping-zip']").val('');
		$('#selected-shipping-address').hide();		
	}	

	$("#shipping-address-container input, #shipping-address-container select").on('change', function(e){
		if($("input[name='shipping-same-billing']").is(':checked'))
			$("input[name='shipping-same-billing']").change();
	});
	
	// When user selects Canada or US IN THE SHIPPING ADDRESS AREA, change divs that display with states and provinces
	$("select[name='shipping-country'], input[name='shipping-city'], select[name='shipping-state'], select[name='shipping-province-canada'], input[name='shipping-province'], input[name='shipping-zip'], input[name='shipping-address'], input[name='shipping-address2']").on('change', function(e){
		
		var varaddress = $("input[name='shipping-address']").val() + " " + $("input[name='shipping-address2']").val();
		var varcity = $("input[name='shipping-city']").val();
		var varstate = $("select[name='shipping-state']").val();
		var varcanada = $("select[name='shipping-province-canada']").val();
		var varprovince = $("input[name='shipping-province']").val();
		var varzip = $("input[name='shipping-zip']").val();
		var varcountry = $("select[name='shipping-country']").val();

		$('#load_temps').load("checkout/inc_store_state_selection.asp", {state:varstate}, function(response, status, xhr) {});

		setCurrency();
		
		if (varcountry == 'USA') {

				//	console.log('CIM ' + var_cim_shipping + ', address ' + cim_address + ', state ' + cim_state);

			var page = $("#ajax-shipping-options");
			page.load("/checkout/ajax_display_shipping_usps.asp", {country:'USA', zip:varzip, address: varaddress, city:varcity, state:varstate}, function(status) {
				$('#shipping-loading').hide();
				$.get("/checkout/ajax_display_shipping_ups.asp", {address: varaddress, country:'USA', city:varcity, state:varstate, zip:varzip, add:'yes'}, function(html, status) {	
					page.append(html);
					triggerShippingSelection();
					calcAllTotals();
				});
			});
				
			$('#shipping-province-canada, #billing-province-canada, #shipping-province, #billing-province').val('');
			$('#shipping-province-canada, #shipping-province').prop('required', false);
			$('#shipping-state').prop('required', true);			
				
			$('.shipping-province-canada').hide();
			$('.billing-province-canada').hide();
			$('.shipping-state').show();
			$('.billing-state').show()
			$('.shipping-province').hide();	
			$('.billing-province').hide();
			$('.hide_inter_zip').hide();
			$('.hide_usa_zip').show();
			$('.customs-notice').hide();
		}
		if (varcountry == 'Canada') {

			$('#ajax-shipping-options').load("/checkout/ajax_display_shipping_usps.asp", {country:'Canada'}, function(){
				triggerShippingSelection();
				calcAllTotals();
			});

			$('#shipping-state, #billing-state, #shipping-province, #billing-province').val('');
			$('#shipping-province-canada').prop('required', true);
			$('#shipping-state, #shipping-province').prop('required', false);
			
			$('.shipping-province-canada, .billing-province-canada, .hide_inter_zip, .customs-notice').show();
			$('.shipping-state').hide();
			$('.billing-state').hide();
			$('.shipping-province').hide();	
			$('.billing-province').hide();
			$('.hide_usa_zip').hide();
		}
		if (varcountry != 'USA' && varcountry != 'Canada') {
			$('#ajax-shipping-options').load("/checkout/ajax_display_shipping_usps.asp", {country:varcountry}, function(){
				triggerShippingSelection();
				calcAllTotals();
			});
					
		//	$('.ups-ajax').load("checkout/ajax_display_shipping_ups.asp", {country:varcountry, city:varcity, province:varprovince, zip:varzip, add:'yes'});

			$('#shipping-state, #billing-state, #shipping-province-canada, #billing-province-canada').val('');
			$('#shipping-state, #shipping-province-canada').prop('required', false);
			$('.shipping-province-canada').hide();
			$('.billing-province-canada').hide();
			$('.shipping-state').hide();
			$('.billing-state').hide();
			$('.shipping-province').show();	
			$('.billing-province').show();
			$('.hide_inter_zip').show();
			$('.hide_usa_zip').hide();	
			$('.customs-notice').show();
		}
		
	//	console.log("State:" + varstate + " Country:" +varcountry+ " Province:" + varprovince)
		verify_shipmethod_selected();
		e.preventDefault();
	}); // End province/ state display or hide
	
	// When user selects Canada or US IN THE BILLING SECTION, change divs that display with states and provinces
	$("select[name='billing-country']").change(function(){
		var varcountry = $("select[name='billing-country']").val();
		
		if (varcountry == 'USA') {			
			$('.billing-province-canada').hide();
			$('.billing-state').show()
			$('.billing-province').hide();
			$('.hide_inter_zip').hide();
			$('.hide_usa_zip').show();
			$('.customs-notice').hide();
		}
		if (varcountry == 'Canada') {
			$('.billing-province-canada').show();
			$('.billing-state').hide();
			$('.billing-province').hide();
			$('.hide_inter_zip').show();
			$('.hide_usa_zip').hide();
			$('.customs-notice').show();			
		}
		if (varcountry != 'USA' && varcountry != 'Canada') {
			$('.billing-province-canada').hide();
			$('.billing-state').hide();
			$('.billing-province').show();
			$('.hide_inter_zip').show();
			$('.hide_usa_zip').hide();	
			$('.customs-notice').show();
		}
	}); // BILLING SECTION End province/ state display or hide	


	// Make sure that only PayPal OR Cash can be checked at once
	$('#paypal').click(function(){
		$('#cash').prop('checked', false);
	});
	$('#cash').click(function(){
		$('#paypal').prop('checked', false);
	});

	// Make sure that only PayPal OR Cash can be checked at once
	$(document).on('click', '.btn-group label, .shipping-method', function() {
		var type = $(this).attr("data-type");
		$('.' + type + ' .btn-selected').html('');
		$('.btn-selected', this).html('<i class="ml-2 fa fa-lg fa-check"></i>');
	});
	
	// Non registered customers, pay by cash
	$('#cash').click(function(){

		if($("#cash").is(':checked')) {
		// Toggle required attribute
		toggleRequiredBillingFalse();
		
		$('.billing-input-fields').fadeOut(1000);
		$('#credit_card_inputs').fadeOut(1000);

		} else {
		// Toggle required attribute
		toggleRequiredBillingTrue();

		$('.billing-input-fields').fadeIn(1000);
		$('#credit_card_inputs').fadeIn(1000);		
		
		}
	});

	// Registered customers pay by cash
	$('#cim_cash_click').click(function(){
		// Toggle required attribute
		toggleRequiredBillingFalse();
		$('.billing-input-fields').fadeOut(1000);
		$('#credit_card_inputs').fadeOut(1000);
	});
	
$('#cardNumber').keyup(function() {	
	//console.log(GetCardType($('#cardNumber').val()));
	$('#cardType').html = GetCardType($('#cardNumber').val())
});

// Function to detect what card type is being used
function GetCardType(number)
{
	if ($('.fa').hasClass('text-primary')) {
		$('.fa').removeClass('text-primary');
	}

    // visa
    var re = new RegExp("^4");
    if (number.match(re) != null) {
		$('.fa-cc-visa').addClass('text-primary');
		return "Visa";
	}

    // Mastercard 
	// Updated for Mastercard 2017 BINs expansion
    var re = new RegExp("^5");
    if (number.match(re) != null) {
		$('.fa-cc-mastercard').addClass('text-primary');
		return "Mastercard";
	}
	/*
	 if (/^(5[1-5][0-9]{14}|2(22[1-9][0-9]{12}|2[3-9][0-9]{13}|[3-6][0-9]{14}|7[0-1][0-9]{13}|720[0-9]{12}))$/.test(number)
	 ) {
		$('.fa-cc-mastercard').addClass('text-primary');
		return "Mastercard";
	 }
	 */

    // AMEX
    re = new RegExp("^3[47]");
    if (number.match(re) != null) {
		$('.fa-cc-amex').addClass('text-primary');
		return "AMEX";
	}

    // Discover
    re = new RegExp("^(6011|622(12[6-9]|1[3-9][0-9]|[2-8][0-9]{2}|9[0-1][0-9]|92[0-5]|64[4-9])|65)");
    if (number.match(re) != null) {
		$('.fa-cc-discover').addClass('text-primary');
		return "Discover";
	}

    return "";
}

// Checkout newsletter signup
$("#checkout-newsletter-signup").on("click", function () {
	if ($("#checkout-newsletter-signup").is(':checked')) {
		$.ajax({
			method: "post",
			dataType: "json",
			url: "/klaviyo/klaviyo-subscribe-newsletter.asp?email=" + $("#e-mail").val()
			})
	} else {
		$.ajax({
			method: "post",
			dataType: "json",
			url: "/klaviyo/klaviyo-subscribe-newsletter.asp?email=" + $("#e-mail").val()
			})		
	}
  });

$(document).on('click', '#btn-edit-shipping-address', function() {
	showShippingAddressInputs();	
});

$(document).on('click', '#btn-edit-billing-address', function() {
	showBillingAddressInputs();
});

function createAddressBubble(section, source) {
	var state = "";
	if (section == 'shipping'){
		state = getShippingState();
		hideShippingAddressValidation();
	}	
	if (section == 'billing'){
		state = getBillingState();
		hideBillingAddressValidation();
	}	
	
	var address = '';	
	if(source == 'fromCookies'){
		address = ($('#' + section + '-address').val() + ' ' + $('#' + section + '-address2').val() + '<br/>') +
		($('#' + section + '-city').val() + '<br/>') +
		($('#' + section + '-country').val() + '<br/>') +
		(state + '<br/>') +
		($('#' + section + '-zip').val() + '');
	}else if(source =='fromShipping'){
		address = $("#shipping-address").val() + ' ' + $("#shipping-address2").val() + '<br />' + 
		$("#shipping-city").val() + '<br />' + 
		state + '<br />' + 
		$("#shipping-zip").val() + '<br />' + 
		$("select[name='shipping-country']").val();
	}	
			
	var content = '<div class="alert alert-secondary alert-dismissible fade show" role="alert">' + 
	'  <div id="selected-' + section + '-address-content" class="m-2"><div class="mb-2 font-weight-bold">' + ((section == 'shipping') ? '<i class="fa fa-shipping-fast fa-lg mr-2"></i> ':'') + '<span style="text-transform: uppercase;">' + section + ' ADDRESS</span></div><div style="line-height:22px;">' + address + '</div></div>' +
	'  <button type="button"  class="close" id="btn-edit-' + section + '-address" style="right:20px;padding: 7px 11px 7px 11px;margin-right:16px">' + 
	'	<img src="/images/edit.svg" style="height:14px;width:14px;vertical-align:initial;" />'  +
	'  </button>' +	
	'  <button id="' + section + '-bubble-close" type="button" class="close" data-dismiss="alert" aria-label="Close" style="padding: 7px 11px 7px 11px" onClick="closeAddressBubble(\'' + section + '\')">' + 
	'	<span aria-hidden="true">&times;</span>' +
	'  </button>' +
	'</div>';
	
    $('#selected-' + section + '-address').html(content);
	$('#selected-' + section + '-address').show();	
	$('#' + section + '-country').change();	
	$("input[name='shipping-address2']").val('');
	$("input[name='billing-address2']").val('');
	
	if($("input[name='shipping-same-billing']").is(':checked') && section == 'shipping')
		$("input[name='shipping-same-billing']").change();
	
}

// If user from the USA or Canada, show the address validation field
$( document ).ready(function() {
	if(isUserFromUSAorCANADA()){
		if($("#shipping-address").val() !="" && $("#shipping-city").val() !="" && $("#shipping-zip").val() !="" && $("#shipping-country").val() !="" && getShippingState() != ""){
			createAddressBubble('shipping', 'fromCookies');		
			$("#shipping-address-container").hide();
		}else{
			clearShippingAddressInputs();
			showShippingAddressValidation();
		}
		showBillingAddressValidation();			
	}else{
		showShippingAddressInputs();
		showBillingAddressInputs();
	}
});


	//If all fields are filled from cookies then create address bubble	
	if($("#shipping-address").val() !="" && $("#shipping-city").val() !="" && $("#shipping-zip").val() !="" && $("#shipping-country").val() !="" && getShippingState() != ""){
		createAddressBubble('shipping', 'fromCookies');		
		$("#shipping-address-container").hide();
	}else{
		showShippingAddressInputs();			
	}