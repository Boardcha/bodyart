// Toggle state / province depending on country selection
$('#country').on('change', function () {

	$('.state, .province, .province-canada').hide();
	$('#state, #province, #province-canada').val('');
	$('#state, #province, #province-canada').prop('required', false);
	
	if ($('#country').val() === "USA") {
		$('#state').prop('required', true);
		$('.state').show();
	} else {			
		
		if ($('#country').val() === "Canada") {
			$('.province-canada').show();
			$('#province-canada').prop('required', true);
		} else {
			$('.province').show();
		}				
	}
	
});

// Load up USA by default on page load. This triggers the state/province fields to show/hide approriately
function resetUSA() {
	$('#country').val('USA').trigger('change');
}

// Input cursor enabled on first field of address modal pop up
$('body').on('shown.bs.modal', '#updateAddress', function () {
    $('input:visible:enabled:first', this).focus();
})

// Function that hides/shows credit card fields whether it's a shipping address or billing address. Also flags them as required if needed.
function requireCardFields() {
	var type = $("#address-type").attr("data-type");

	if (type == 'shipping') {
		$('#cardNumber, #creditCardMonth, #creditCardYear').prop('required', false);
		$('.address-fields').show();
		$('.card-fields').hide();
	}
	if (type == 'billing') {
		$('#cardNumber, #creditCardMonth, #creditCardYear').prop('required', true);
		$('.address-fields, .card-fields').show();
	}
}

// Change form submit status depending on what icon/button is being clicked
$('#add-address, .btn-edit').click(function () {
	requireCardFields();
	$('#frm-cim')[0].reset();
	resetUSA();

	$('#frm-cim').removeClass('was-validated');
	$('.message-address-modal').hide();

	var type = $("#address-type").attr("data-type");
	var status = $(this).attr("data-status");
	var url = $(this).attr("data-url");
	var addressId = $(this).attr("data-id");
	var header = $(this).attr("data-header");
	var buttonText = $(this).attr("data-buttonText");
	$("#frm-cim").attr("data-type",type);
	$("#type").val(type);
	$("#frm-cim").attr("data-status",status);
	$('#cim-id').val(addressId);
	$("#frm-cim").attr("data-url",url);	
	$("#headerAddress").html(header);
	$("#btn-modal-address").html(buttonText);

});

// Make an address the default address
$('.make-default').click(function () {
	var id = $(this).attr("data-id");
	var type = $("#address-type").attr("data-type");
	
	$.ajax({
	method: "post",
	url: "accounts/ajax-cim-make-default.asp",
	data: {id: id, type: type},
	context: this
	})
	.done(function(msg) {	
		$(this).html('<h6 class="text-dark d-inline">DEFAULT ADDRESS<h6>');
		$("#default-address").hide();
	})
	.fail(function(msg) {
		$(this).html('<span class="alert alert-danger">ERROR</span>');
	})
});  // END make an address the default address


// Display selected address in modal window for confirmation
$('.delete-address').click(function (e) {
	var id = $(this).attr("data-id");
	$('#confirm-delete').attr("data-id",id);
	var address = $(this).attr("data-address");
	$("#modal-confirm-address").html(address);
	
});

	
// Delete shipping or billing profiles
	$('#confirm-delete').click(function (e) {
		var type = $("#address-type").attr("data-type");
		var id = $(this).attr("data-id");
		$("#confirm-delete").attr("data-type",type);

		$.ajax({
		method: "post",
		dataType: "json",
		url: "accounts/ajax-cim-delete-address.asp",
		data: {id: id, type: type}
		})
		.done(function(json, msg) {
			if (json.status == 'success') {
				// Close modal window
				$('#deleteModal').modal('toggle');
				window.scrollTo(0,0);
				$('.message-window').html('<div class="alert alert-success">Your address has been deleted.</div>').show();
				$('.message-window').delay(5000).fadeOut('slow');
				$('.' + id + '-block').fadeOut('slow');
				$('.' + id + '-block').removeClass("d-inline-block");

				
			}
			
		})
		.fail(function(json, msg) {
			$('.message-delete-modal').html('<div class="notice-red">Error. Please make sure you entered a valid address. If you continue to have trouble, please contact customer service for assistance.</div>').show();
		})
	});  // END Delete shipping or billing profiles
	
	
	// START Pre-populate and bring up edit CIM form
	$('.btn-edit').click(function () {
		requireCardFields();
		
		var first = $(this).attr("data-first");
		var last = $(this).attr("data-last");
		var company = $(this).attr("data-company");
		var address = $(this).attr("data-address");
		var address2 = $(this).attr("data-address2");
		var city = $(this).attr("data-city");
		var state = $(this).attr("data-state");
		var zip = $(this).attr("data-zip");
		var country = $(this).attr("data-country");
		
		$('#first').val(first);
		$('#last').val(last);
		$('#company').val(company);
		$('#address').val(address);
		$('#address2').val(address2);
		$('#city').val(city);		
		$('#zip').val(zip);
		$('#country').val(country);
		$("#country").trigger('change');
		if ($('#country').val() === "USA") {
			$('#state').val(state);
		} else {			
			if ($('#country').val() === "Canada") {
				$('#province-canada').val(state);
			} else {
				$('#province').val(state);
			}				
		}
	});  // END Pre-populate and bring up edit CIM form 

	// On form submit dynamically add or update an address
	$('#frm-cim').submit(function (e) {

		var url = $(this).attr("data-url");
		var type = $("#address-type").attr("data-type");
		var status = $(this).attr("data-status");
		var addressId = $('#cim-id').val();

		// Used to display new card or updated card information
		var first = $('#first').val();
		var last = $('#last').val();
		var company = $('#company').val();
		var address = $('#address').val();
		var address2 = $('#address2').val();
		var city = $('#city').val();
		var state = $('#state').val();
		var province = $('#province').val();
		var province_canada = $('#province-canada').val();
		var zip = $('#zip').val();
		var country = $('#country').val();
		var cc_num = $('#cardNumber').val();
		var lastFour = cc_num.substr(cc_num.length - 4);
		var cc_month = $('#billing-month').val();
		var cc_year = $('#billing-year').val();

		// Fetch form to apply custom Bootstrap validation
		var form = $("#frm-cim")

		if (form[0].checkValidity() === false) {
			e.preventDefault()
			e.stopPropagation()
			console.log("invalid form elements");
		} else {
		
		$.ajax({
		method: "post",
		dataType: "json",
		url: "accounts/" + url,
		data: $("#frm-cim").serialize()
		})
		.done(function(json, msg, data) {
			if(json.status == "success") {
				// Close modal window
				$('#updateAddress').modal('toggle');
				window.scrollTo(0,0);
				$('.message-window').html('<div class="alert alert-success">Your ' + type + ' information has been successfully updated.</div>').show();
				$('.message-window').delay(5000).fadeOut('slow');

				if(status === 'add') { // If address being added then show new card
				if (type == 'shipping') {
				// Display address info
				$('#show-new').prepend('<div class="card d-inline-block m-md-2 my-2 account-cards"><div class="card-body bg-light">' + first + ' ' + last + '<br/>' + company + '<br/>' + address + ' ' + address2 + '<br/>' + city + ', ' + state + province + province_canada + '  ' + zip + '</div></div>');
				
				} else {
					// Display card info
				$('#show-new').prepend('<div class="card d-inline-block m-md-2 my-2 account-cards"><div class="card-body bg-light"><div class="font-weight-bold"> ' + lastFour + '</div>' + address + ' ' + address2 + '<br/>' + city + ', ' + state + province + province_canada + '  ' + zip + '</div></div>');
				}
			} else { // if address is being updated, then update current card or address

					// Update edit button attributes in case they need to edit again
					$('.edit-' + addressId).attr("data-first",first);
					$('.edit-' + addressId).attr("data-last",last);
					$('.edit-' + addressId).attr("data-company",company);
					$('.edit-' + addressId).attr("data-address",address);
					$('.edit-' + addressId).attr("data-address2",address2);
					$('.edit-' + addressId).attr("data-city",city);
					$('.edit-' + addressId).attr("data-state",state + province + province_canada);
					$('.edit-' + addressId).attr("data-zip",zip);
					$('.edit-' + addressId).attr("data-country",country);

				if (type == 'shipping') {
					
					// Display address info
					$('.' + addressId + '-address').html(first + ' ' + last + '<br/>' + company + '<br/>' + address + ' ' + address2 + '<br/>' + city + ', ' + state + province + province_canada + '  ' + zip + '<br/>' + country);
					} else {
						// Display card info
					// Display address info
					$('.' + addressId + '-address').html('<div class="font-weight-bold">Card ending in ' + lastFour + '</div>' + first + ' ' + last + '<br/>' + company + '<br/>' + address + ' ' + address2 + '<br/>' + city + ', ' + state + province + province_canada + '  ' + zip + '<br/>' + country);				
					}
			}

				
				
			} else {
				$('.message-address-modal').html('<div class="alert alert-danger">Error. Please make sure you entered a valid address. <span class="font-weight-bold">Detailed message: ' + json.message + '</span></div>').show();
			}			
		})
		.fail(function(json, msg) {
			$('.message-address-modal').html('<div class="alert alert-danger">' + json.message + ' Website error. If you continue to have trouble, please contact customer service for assistance.</div>').show();
		})
		}
		form[0].classList.add('was-validated');
		e.preventDefault();

	}); // End On form submit