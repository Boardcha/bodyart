// HIde fields on page load 
$('#cardNumber, #creditCardMonth, #creditCardYear').prop('required', false);
$('.card-fields').hide();


	// Load tracking information
	$(document).on('click', '.track', function() {
		var number = $(this).attr('data-num');
		var invoice = $(this).attr('data-invoice');
		var url = $(this).attr("data-url");
		$('#track-body').load(url + number);	
	});

	// Change out modal attributes when clicking buttons. Keeping it lazy by just changing attributes on all forms even though it's not correct until they press a button.
	$(document).on('click', '.btn-update-attributes', function() {
		$('.modal-message').html('');
		$('.modal-submit').show();
		var orderitemid = $(this).attr("data-orderitemid");
		var title = $(this).attr('data-title');

		$('form').attr("data-orderitemid",orderitemid);
		$('.title').html(title);	
	});
	
	$(document).on('click', '.btn-report-problem', function() {
        var id = $(this).attr('data-invoice');
		$('#message-problem-modal').html('');
		$('#confirm-report-problem').hide();
        $('#loader-report-problem').load("/accounts/ajax-report-order-problem.asp", {id: id}); 
    });	
 
    // Hide all other items once a reported error item has been selected. This will save space in the modal window for smaller screens.
	$(document).on('click', '.problem-select-item', function() {
		// When choosing an item on the report order problem, change the qty in the how many missing field to be the max of what they ordered
		var qty = $(this).attr('data-qty');	
		$('#qty-missing').val(qty);
		
		$('.problem-select-item').not($(this)).hide();
		$('#block-select-problem, #confirm-report-problem').show();
		$('#header-select-item').hide();
		$(this).addClass('alert alert-info pl-5 py-1');
    });	

	// Toggle how many missing field on report order problem
	$(document).on('change', 'input[name="status"]', function() {
		var status_value = $(this).val();
		if (status_value == 'Missing' || status_value == 'Broken' || status_value == 'Wrong') {
			$('.qty-missing').fadeIn();
			$('#error-type').html(status_value);
		} else {
			$('.qty-missing').fadeOut();
			$('#error-type').html('');
		}
	});		
	
	// START submit order problem
	$(document).on('click', '#confirm-report-problem', function(event) {
		$.ajax({
				method: "post",
				url: "accounts/ajax-report-order-problem-submit.asp",
				data: $('#form-report-problem').serialize()
				})
				.done(function( msg) {
					$('#loader-report-problem').html('');
					$('#confirm-report-problem').hide();
					$('#message-problem-modal').html('<div class="alert alert-success p-1 my-2">Thank you for letting us know about your order problem.<p>Please give 1-2 working days for your request to be reviewed by our customer service department. If we have any questions about your request/order problem, we will contact you via e-mail. If you receive a shipment notice in the next 1-2 working days, then it is your item(s) being shipped out (please do not be alarmed).</p></div>').show();
				})
				.fail(function(msg) {
					$('#message-problem-modal').html('<div class="alert alert-danger p-1 my-2">Website error. Please check that all fields are filled out. If you continue to have trouble, please contact customer service for assistance.</div>').show();
				})	
	});	// END submit order problem

	// START load more information item button
	$(document).on('click', '.btn-moreInfo-modal', function() {
		var preorder_specs = $(this).attr('data-preorderSpecs');
		
		if (preorder_specs != '') {
			preorder_specs = '<div class="font-weight-bold mb-1">Your Pre-Order Specifications:</div>' + preorder_specs
		}
console.log(preorder_specs);
		$('#loader-more-information').html(preorder_specs)
	});

	// Set default star values because being a hidden radio buttin it's being wiped onLoad.
	function setStarValues() {
		$('.star1').val('1');
		$('.star2').val('2');
		$('.star3').val('3');
		$('.star4').val('4');
		$('.star5').val('5');
		};
		setStarValues();

	// START show/hide info on modal to review a product
	$(document).on('click', '.btn-review-modal', function(e) {
		$('#review_order_detail_id').val($(this).attr('data-orderitemid'));
		$('#review-body').show();
		$('#text-review').val('');
		setStarValues();
		$('#frm-write-review [name="rating"]').prop('checked', false);
	});

	// START submit item review
	$(document).on('submit', '#frm-write-review', function(e) {
		var order_detail_id = $('#review_order_detail_id').val();
		var review = $("#frm-write-review [name='review']").val();
		var rating = $("#frm-write-review [name='rating']:checked").val();

		// Fetch form to apply custom Bootstrap validation
		var form = $("#frm-write-review")

		if (form[0].checkValidity() === false) {
			e.preventDefault()
			e.stopPropagation()
			console.log("invalid form elements");
		} else {


		$.ajax({
			method: "post",
			url: "products/ajax-write-product-review-submit.asp",
			contentType: "application/x-www-form-urlencoded;charset=UTF-8",
			data: {id: order_detail_id, review: review, rating: rating}
			})
			.done(function( msg) {
				$('#message-review-modal').html('<div class="alert alert-success">Thank you for reviewing this product. Once your review is accepted and published, you will see the points reflected on your account that you can exchange for store credit.</div>');
				
				$('.modal-submit').hide();
				$('.review-phase1-' + order_detail_id).removeClass('btn-update-attributes');
				$('.review-phase1-' + order_detail_id).addClass('alert-success');
				$('.review-phase1-' + order_detail_id).removeAttr('data-title data-orderitemid data-toggle data-target');
			})
			.fail(function(msg) {
				$('#message-review-modal').html('<div class="alert alert-danger">Website error. Please check that your review is filled out. If you continue to have trouble, please contact customer service for assistance.</div>');
			})
		}
		form[0].classList.add('was-validated');
		e.preventDefault();	

	});	// END submit item review
	

	// START show/hide info on modal to review a product
	$(document).on('click', '.btn-photo-modal', function(e) {
		$('#photo-id').val($(this).attr('data-productid'));
		$('#photo-orderdetailid').val($(this).attr('data-orderitemid'));
		$('#photo-detailid').val($(this).attr('data-detailid'));
		$('#photo-email').val($('#copy-photo-email').val());
		$('#photo-name').val($('#copy-photo-name').val());
		$('#photo-custid').val($('#copy-photo-custid').val());
		$('#message-photo-modal').html('');
		$('#photo-filename').show();
	});

	// START submit PHOTO 
	$(document).on('click', '#confirm-submit-photo', function (event) {
		var order_detail_id = $('#photo-orderdetailid').val();	
		var myForm = document.getElementById('frm-submit-photo');
		var form_data = new FormData($(myForm)[0]);
		$('#photo-filename').hide();
		$('#message-photo-modal').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');
		
		// Submit photo
		$.ajax({
		method: "post",
		url: "/gallery/ajax-upload-photo-version1.asp",
		data: form_data,
		dataType: "json",
	//	cache: false,
		processData: false,
		contentType: false
		})
		.done(function(json, data, msg) {
		
			if (json.status == "success") {
				$('#message-photo-modal').html('<div class="alert alert-success">' + json.status_text + '</div>');
						
				$('.modal-submit').hide();
				$('.photo-phase1-' + order_detail_id).removeClass('btn-update-attributes');
				$('.photo-phase1-' + order_detail_id).addClass('alert-success');
				$('.photo-phase1-' + order_detail_id).removeAttr('data-title data-orderitemid data-toggle data-target');
			}
			
			if (json.status == "duplicate") {
				$('#message-photo-modal').html('<div class="alert alert-danger p-1 my-2">' + json.status_text + '</div>')
				$('#photo-filename').show();
				console.log(order_detail_id);
				console.log(json.order_id);
			}
			
		})
		.fail(function(json, msg) {
			$('#message-photo-modal').html('<div class="alert alert-danger p-1 my-2">Photo failed to upload. Problems could be: duplicate file name or a file size larger than 1.5MB. Try cropping your photo down a bit and see if it works.</div>')
			$('#photo-filename').show();
		})
	
	});  // END submit PHOTO

	

// Bring up cancel modal with certain displayed items
$(document).on('click', '.btn-cancel-modal', function() {
	var id = $(this).attr('data-invoice');
	$('#confirm-cancel').attr("data-invoice", id); // Update attribute with invoice
	$('#message-cancel-modal').html('');
	$('#loader-cancel-modal').load("/accounts/ajax-cancel-order.asp", {id: id});

	$.ajax({
		method: "post",
		dataType: "json",
		url: "/accounts/ajax-cancel-double-check.asp",
		data: {id: id}
		})
		.done(function(json, msg) { 
				if(json.status === "success") { 
					$('#confirm-cancel').show();
				} else {
					$('#confirm-cancel').hide();
				}
		})
	});	


	
// START confirm order cancellation
	$(document).on('click', '#confirm-cancel', function(event) {
		var invoice = $(this).attr('data-invoice');
		$.ajax({
		method: "post",
		url: "accounts/ajax-cancel-order-submit.asp",
		data: {invoice:invoice}
		})
		.done(function(msg) {
			$('#loader-cancel-modal').html('');	
			$('#confirm-cancel, #cancel-' + invoice).hide();		
			$('#message-cancel-modal').html('<div class="alert alert-success p-1 my-2">Your order has been cancelled. Your store credit is now on your account to be used..</div>').show();
		})
		.fail(function(msg) {
			$('#message-cancel-modal').html('<div class="alert alert-danger p-1 my-2">Website error. There was a problem cancelling your order. If you continue to have trouble, please contact customer service for assistance.</div>').show();
		})
	});  // END confirm order cancellation

	// Open survey modal to transfer id #
	$(document).on('click', '.btn-survey-modal', function() {
		var id = $(this).attr('data-id');
		$('#surveyId').val(id);
		$('#message-survey-modal').html('');
		$('#frm-survey').show();
		$('#frm-survey')[0].reset();
		$('.wrapper-textarea').hide();
		$("#survey-body").animate({ scrollTop: 0 }, "fast");
	});		
	
	// Order survey toggle textarea based on star rating value
	$(document).on('click', '#frm-survey label', function() {
		var group = $(this).attr('data-group');
		var answer = $(this).attr('data-value');
		var radioid = $(this).attr('for');
		

		// Check radio button based on star selection
		if (answer < 4) {
			$('.textarea-' + group).fadeIn();
			$('#' + group + 'Elaborate').prop('required', true);			
		} else {
			$('.textarea-' + group).fadeOut();
			$('#' + group + 'Elaborate').prop('required', false);
		}
	});	
	
// START submit order survey
	$(document).on('click', '#confirm-submit-survey', function(e) {	
		var form = $("#frm-survey")

        if (form[0].checkValidity() === false) {
            e.preventDefault()
            e.stopPropagation()
			console.log("invalid form elements");
			$('#message-survey-modal').html('<div class="alert alert-danger p-1 my-2">Some ratings have not been selected. Please scroll up and make sure all ratings (highlighted in red) have been selected.</div>').show();
        } else {

		$.ajax({
		method: "post",
		url: "accounts/ajax-order-survey-submit.asp",
		data: $("#frm-survey").serialize()
		})
		.done(function(msg) {			
			$('#message-survey-modal').html('<div class="alert alert-success p-1 my-2">Thanks! Your survey has been submitted and the .50¢ store credit is on your account.</div>').show();
		//	$('#frm-survey').hide();			
		})
		.fail(function(msg) {
			$('#message-survey-modal').html('<div class="alert alert-danger p-1 my-2">Website error. Please <a class="alert-link" href="/contact.asp">contact customer service</a> for assistance.</div>').show();
		})
		}
		form[0].classList.add('was-validated');
		e.preventDefault();
	});  // END submit order survey

// On update address open modal, set form to USA fields, and scroll window to top, and set invoice ID for submit button
$(document).on('click', '.btn-update-address-modal', function() {
	var id = $(this).attr('data-invoice');
	var country = $(this).attr('data-country');
	$('#message-address-modal').html('');
	$('#country').val(country).trigger('change');
	$('#confirm-update-address').attr("data-id",id);
	$("#update-address-body").animate({ scrollTop: 0 }, "fast");
});
	
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

// On form submit dynamically add or update an address
$(document).on('click', '#confirm-update-address', function(e) {	
	// Fetch form to apply custom Bootstrap validation
	var form = $("#frm-update-address")
	var id = $(this).attr('data-id');

	if (form[0].checkValidity() === false) {
		e.preventDefault()
		e.stopPropagation()
		console.log("invalid form elements");
	} else {

	$.ajax({
	method: "post",
	dataType: "json",
	url: "accounts/ajax-update-order-address.asp",
	data: $("#frm-update-address").serialize() + "&id=" + id
	})
	.done(function(json, msg, data) {
		if(json.status == "success") {
			$('#message-address-modal').html('<div class="alert alert-success p-1 my-2">Your shipping address has been successfully updated.</div>').show();		
			$('#address-box-' + id).html($('#company').val() + '</br>' + $('#first').val() + ' ' + $('#last').val() + '</br>' + $('#address').val() + ' ' + $('#address2').val()+ '</br>' + $('#city').val() + ', ' + $('#state').val() + $('#province-canada').val() + ' ' + $('#province').val() + '</br>' + $('#country').val());
		} else {
			$('#message-address-modal').html('<div class="alert alert-danger p-1 my-2">Error. Please make sure you entered a valid address. If you continue to have trouble, please contact customer service for assistance.</div>').show();
		}			
	})
	.fail(function(json, msg) {
		$('#message-address-modal').html('<div class="alert alert-danger p-1 my-2">Website error. If you continue to have trouble, please contact customer service for assistance.</div>').show();
	})
	}
	form[0].classList.add('was-validated');
	e.preventDefault();

}); // End On form submit


// Bring up Add-On modal with certain displayed items
$(document).on('click', '.btn-addon-modal', function() {
	var id = $(this).attr('data-invoice');
	$('#confirm-start-addons').attr("data-invoice", id); // Update attribute with invoice
	$('#confirm-start-addons').show();
});	

// START confirm order add-ons
$(document).on('click', '#confirm-start-addons', function(event) {
	var invoice = $(this).attr('data-invoice');
	Cookies.set('OrderAddonsActive', invoice, { expires: 30});
	window.location = "/products.asp?new=Yes";
});  // END confirm order add-ons

