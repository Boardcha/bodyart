	var totalWithoutShipping = 0;
	var salesTax = 0;
	var shippingWeight = 0;
	var free_items_count = 0;
	
	function calcAllTotals(e) {
		
		$('.checkout_button, #paypal-button-container').hide();
		var shipping_option = $('input[name="shipping-option"]:checked').val();
		$(".cart_grand-total").html('<i class="fa fa-spinner fa-spin"></i> Calculating');
		
		var page_name =  location.pathname.substring(location.pathname.lastIndexOf("/") + 1);

		// Tax calculation calls TAXJAR
		var tax_country = "";
		var tax_state = "";
		var tax_zip = "";
		var tax_address = "";
		var state_taxed = "";

		if ($("input:radio[name='cim_shipping']").is(':checked')) {
			tax_country = $("input[name='cim_shipping']:checked").attr("data-country");
			tax_state = $("input[name='cim_shipping']:checked").attr("data-state");
			tax_zip = $("input[name='cim_shipping']:checked").attr("data-zip");
			tax_address = $("input[name='cim_shipping']:checked").attr("data-address");
		} else {
			tax_country = $("select[name='shipping-country']").val();
			tax_state = $("select[name='shipping-state']").val();
			tax_zip = $("input[name='shipping-zip']").val();
			tax_address = $("input[name='shipping-address']").val() + " " + $("input[name='shipping-address2']").val();

		}
		if (tax_state === 'AR' || tax_state === 'CA' || tax_state === 'CO' || tax_state === 'FL' || tax_state === 'GA' || tax_state === 'HI' || tax_state === 'IL' || tax_state === 'IN' || tax_state === 'IA' || tax_state === 'KS' || tax_state === 'KY' || tax_state === 'LA' || tax_state === 'MA' || tax_state === 'ME' || tax_state === 'MD' || tax_state === 'MI' || tax_state === 'MN' || tax_state === 'NE' || tax_state === 'NV' || tax_state === 'NJ' || tax_state === 'NC' || tax_state === 'OH' || tax_state === 'OK' || tax_state === 'PA' || tax_state === 'RI' || tax_state === 'SD' || tax_state === 'TX' || tax_state === 'UT' || tax_state === 'VA' || tax_state === 'VT' || tax_state === 'WA' || tax_state === 'WV') {
			state_taxed = "yes";
		}
		// END Tax calculation calls

		$.ajax({
			type: "post",
			async: false,
			url: "cart/inc_cart_totals.asp",
			dataType : "json",
			data: {shipping_option: shipping_option, page_name: page_name, tax_country: tax_country, tax_state: tax_state, tax_zip: tax_zip, tax_address: tax_address, state_taxed: state_taxed},
			success: function( json ) {
				//console.log("state taxes: " + state_taxed + " in " + tax_state);
				
				
			subTotal = parseFloat(json.subtotal);
			totalDiscount = parseFloat(json.total_discount);
			salesTax = json.salestax;
			$(".cart_subtotal").html(json.subtotal);
			$(".cart_grand-total").html(json.grandtotal);
			$('.convert-total').attr('data-price',json.grandtotal);
			$(".cart_sales-tax").html(json.salestax);
			$(".cart_coupon-amt").html(json.couponamt);
			$(".cart_prefferred_discount").html(json.preferred_discount);
			$(".shipping_amount_left").html(json.shippingneeded);
			$(".cart_shipping").html(json.shippingfriendly);
			$("#cart_gift-cert").html(json.var_total_giftcert_used);
			$("#store_credit_amt").html(json.store_credit_amt);
			$("#use_now_amount").html(json.use_now_credits);
			shippingWeight = json.weight;
			preOrderItem = json.pre_order_item;

			if (tax_country === 'USA') {
				$(".cart_sales-tax-state").html(json.salestax_state + ' ');
			}
			/*
			if (tax_country === 'Great Britain' || tax_country === 'Great Britain and Northern Ireland' || tax_country === 'United Kingdom') {
				$(".cart_sales-tax-state").html('VAT ');
			} */

			//AfterpayWidget();
				
			// Update all the currencies
			if (Cookies.get("currency") != 'USD') {
				updateCurrency();
			} else {
				$('.row_convert_total').hide();
			}
			
				if (typeof shipping_option != "undefined") {
					$('.checkout_button').show();
				} else {
					// 
					if($('#gift_cert_only').val() === 'yes') {
						$('.checkout_button, #paypal-button-container').show();
					}	
				}
			if (page_name === "cart.asp" || page_name === "cart2.asp") {
					$('.checkout_button').show();
					// Show AfterPay if total is over $100
					if (json.grandtotal >= 100) {
						$('#btn-afterpay-checkout, .checkout_afterpay').show();	
					}

					if (json.grandtotal <= 0) { // to prevent store credits with $0 balance from using paypal checkout
						$('.payment-options').html('STORE CREDIT');
						$('.checkout_paypal, #btn-afterpay-checkout, .checkout_afterpay, #btn-googlepay, #btn-applepay').hide();
					}	
				}
								
				// Show/hide use now order credits
				if (json.use_now_credits == '$0.00') {
					$('#row_use_now_credits').hide();
				} else {
					$('#row_use_now_credits').show();
				}	
				
				// Show/hide gift cert
				if (json.var_total_giftcert_used == "$0.00") {
					$('#row_gift_cert').hide();
				} else {
					$('#row_gift_cert').show();
				}	

				// Show/hide store credit
				if (json.store_credit_amt == "$0.00") {
					$('#row_store_credit').hide();
				} else {
					$('#row_store_credit').show();
				}	
				
				// Show/hide billing section based on total
				if(!$("#cim_cash, #cash").is(':checked')) {
					if (json.grandtotal <= 0) {
						$('.billing-information').fadeOut(500);
						$('#cash').prop('checked', false);
						$('#paypal').prop('checked', false);
						$('input[name="cim_billing"]:first').trigger("click");
						toggleRequiredBillingFalse();
					} else if ($('#paypal-checkout').val() === 'on') {
						$('.billing-information').hide();
						$('.billing-information :input').prop('disabled', true);
					}					
					else {
						$('.billing-information').fadeIn(1000);
						toggleRequiredBillingTrue();
					}
				}
			
			// Only run this code if cart has items other than gift certs		
			if(json.var_other_items == 1) {
				
				// show basic free items
				Cookies.set('orings', '', { expires: 30});
			
				// Show/hide amount needed for discounted shipping
				if(json.subtotal_after_discounts <= 50) {
					$('.cart_shipping_amountLeft').show();
				} 
				else {
					$('.cart_shipping_amountLeft').hide();
				}
				
				var subtotal_free_gifts = (json.fraudcheck_freegifts_subtotal);
				// removed from line above  - json.var_totalvalue_certs_incart
				
				free_items_count = 7;
				freeGiftText = '';
				
				if(subtotal_free_gifts < 150) {
					free_items_count = 6;
					amountToGetAnotherGift = 150 - subtotal_free_gifts;
					freeGiftText = '<br/><span class="font-weight-normal">You’ve <b>$' + amountToGetAnotherGift.toFixed(2) + '</b> to go to get another free gift.<br><a href="#" data-toggle="modal" data-target="#free-items-page-modal">View our free gifts</a></span>';
					clearFreeItemsCookie(7);
				}
				if(subtotal_free_gifts < 100) {
					free_items_count = 5;
					amountToGetAnotherGift = 100 - subtotal_free_gifts;
					freeGiftText = '<br/><span class="font-weight-normal">You’ve <b>$' + amountToGetAnotherGift.toFixed(2) + '</b> to go to get another free gift.<br><a href="#" data-toggle="modal" data-target="#free-items-page-modal">View our free gifts</a></span>';
					clearFreeItemsCookie(6);
				}
				if(subtotal_free_gifts < 75) {
					free_items_count = 4;
					amountToGetAnotherGift = 75 - subtotal_free_gifts;
					freeGiftText = '<br/><span class="font-weight-normal">You’ve <b>$' + amountToGetAnotherGift.toFixed(2) + '</b> to go to get another free gift.<br><a href="#" data-toggle="modal" data-target="#free-items-page-modal">View our free gifts</a></span>';
					clearFreeItemsCookie(5);
				}
				if(subtotal_free_gifts < 50) {
					free_items_count = 3;
					amountToGetAnotherGift = 50 - subtotal_free_gifts;
					freeGiftText = '<br/><span class="font-weight-normal">You’ve <b>$' + amountToGetAnotherGift.toFixed(2) + '</b> to go to get another free gift.<br><a href="#" data-toggle="modal" data-target="#free-items-page-modal">View our free gifts</a></span>';
					clearFreeItemsCookie(4);
				}

				if(subtotal_free_gifts < 30) {
					free_items_count = 2;				
					amountToGetAnotherGift = 30 - subtotal_free_gifts;
					freeGiftText = '<br/><span class="font-weight-normal">You’ve <b>$' + amountToGetAnotherGift.toFixed(2) + '</b> to go to earn a free gift.<br><a href="#" data-toggle="modal" data-target="#free-items-page-modal">View our free gifts</a></span>';
					clearFreeItemsCookie(3);
				}
				
				if(free_items_count > 0){
					let s = (free_items_count > 1) ? 's' : '';
					$("#free-items-info").html('<span class="font-weight-bold"><i class="fa fa-gift"></i> You’ve earned ' + free_items_count + ' FREE GIFT' + s + ' <a href="#" id="btn-free-items" data-toggle="modal" data-target="#free-items-modal">Select it now</a></span>' + freeGiftText);
				}	
			}else{
				free_items_count = 0;		
				$("#free-items-info").html('');
				clearFreeItemsCookie(1);
			}
			console.log(free_items_count);
			// Disabling place order button by country or other restrictions
			/*	
			if (tax_country === 'Hong Kong' || tax_country === 'Slovenia') {
				$('.checkout_button').hide();
				$('.processing-message').html('<div class="alert alert-danger mt-2 font-weight-bold">CORONAVIRUS (COVID-19) NOTICE<div class="small mt-2">Unfortunately, shipments to your country are temporarily suspended due to the Coronavirus. We will resume shipments when we get the notice from our couriers that it is safe to do so.</div></div>').show();
			} else {
				$('.checkout_button').show();
				$('.processing-message').html('').hide();
			}
			*/

		}
		});	
	}