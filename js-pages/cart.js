$('.checkout_now, .checkout_paypal, #btn-googlepay, #btn-applepay').show();		

	// Remove discounts
	$("#remove-credit").click(function() {
		
		var remove_type = $(this).attr("data-type");		
		$.ajax({
		method: "POST",
		dataType: "json",
		url: "cart/ajax_remove_discounts.asp",
		data: {remove_type: remove_type}
		})
		.done(function( json, msg ) {
			window.location = "/cart.asp";	
		})
		.fail(function(json, msg) {
			console.log("fail");
		});		
	});	 
 
 // Image combo box
	 $(document).ready(function(e) {
	try {
	$("body [name=freegift1],[name=freegift2],[name=freegift3],[name=freegift4],[name=freegift5],[name=freesticker],[name=free_orings]").msDropDown();
	} catch(e) {
//	console.log(e.message);
	}
	});




	// Show proceed to checkout button -- enabled for people with scripts on
	$("#btn_continue_checkout").show();
	
	
	$('.remove_gaugeCard').click(function(){
		$('.free_gauge_card').css('backgroundColor','#F5A9A9');
		$('.free_gauge_card').fadeOut( "slow" );
		Cookies.set('gaugecard', 'no', { expires: 30});
		});
		
	$('.remove_sticker').click(function(){
		$('.free_sticker').css('backgroundColor','#F5A9A9');
		$('.free_sticker').fadeOut( "slow" );
		Cookies.set('sticker', 'no', { expires: 30});
		});
		
	// START Set cookies for free gift selections ------------------------
		function ToggleDisplayUseNow(that, name) {
			var_freegift_text = $(that).find(':selected').data("text");
			var_usenow_notice = name;
			
				if(var_freegift_text.indexOf("USE NOW") != -1){
					$('#' + var_usenow_notice).show();
				} else {
					$('#' + var_usenow_notice).hide();
				}

		}

		$('input[name=freegift1]').change(function(){
			var_freegift1id = $(this).val();
			friendly_title = $(this).attr('data-friendly');
			img_name = $(this).attr('data-img-name');
			Cookies.set('freegift1id', var_freegift1id, { expires: 30});
			$('#selected-gift1').html('<img class="ml-1 mr-2" src="https://s3.amazonaws.com/bodyartforms-products/' + img_name + '"   style="width:40px">' + friendly_title);
			$('#gift1-dropdown-text').hide();
			calcAllTotals();
		});
		
		$('input[name=freegift2]').change(function(){
			var_freegift2id = $(this).val();
			friendly_title = $(this).attr('data-friendly');
			img_name = $(this).attr('data-img-name');
			Cookies.set('freegift2id', var_freegift2id, { expires: 30});
			$('#selected-gift2').html('<img class="ml-1 mr-2" src="https://s3.amazonaws.com/bodyartforms-products/' + img_name + '"   style="width:40px">' + friendly_title);
			$('#gift2-dropdown-text').hide();
			calcAllTotals();
		});
		
		$('input[name=freegift3]').change(function(){
			var_freegift3id = $(this).val();
			friendly_title = $(this).attr('data-friendly');
			img_name = $(this).attr('data-img-name');
			Cookies.set('freegift3id', var_freegift3id, { expires: 30});
			$('#selected-gift3').html('<img class="ml-1 mr-2" src="https://s3.amazonaws.com/bodyartforms-products/' + img_name + '"   style="width:40px">' + friendly_title);
			$('#gift3-dropdown-text').hide();
			calcAllTotals();
		});
		
		$('input[name=freegift4]').change(function(){
			var_freegift4id = $(this).val();
			friendly_title = $(this).attr('data-friendly');
			img_name = $(this).attr('data-img-name');
			Cookies.set('freegift4id', var_freegift4id, { expires: 30});
			$('#selected-gift4').html('<img class="ml-1 mr-2" src="https://s3.amazonaws.com/bodyartforms-products/' + img_name + '"   style="width:40px">' + friendly_title);
			$('#gift4-dropdown-text').hide();
			calcAllTotals();
		});
		
		$('input[name=freegift5]').change(function(){
			var_freegift5id = $(this).val();
			friendly_title = $(this).attr('data-friendly');
			img_name = $(this).attr('data-img-name');
			Cookies.set('freegift5id', var_freegift5id, { expires: 30});
			$('#selected-gift5').html('<img class="ml-1 mr-2" src="https://s3.amazonaws.com/bodyartforms-products/' + img_name + '"   style="width:40px">' + friendly_title);
			$('#gift5-dropdown-text').hide();
			calcAllTotals();
		});
		
	// END Set cookies for free gift selections --------------------------
		
	// Set cookies for free sticker selection
	$('input[name=freesticker]').change(function(){
		var_stickerid = $(this).val();
		friendly_title = $(this).attr('data-friendly');
		img_name = $(this).attr('data-img-name');
		Cookies.set('stickerid', var_stickerid, { expires: 30});
		$('#selected-sticker').html('<img class="ml-1 mr-2" src="https://s3.amazonaws.com/bodyartforms-products/' + img_name + '"   style="width:40px">' + friendly_title);
		$('#sticker-dropdown-text').hide();
		});
		
	// Set cookies for free o-rings selection
	$('input[name=free_orings]').change(function(){
		var_oringsid = $(this).val();
		Cookies.set('oringsid', var_oringsid, { expires: 30});
		friendly_title = $(this).attr('data-friendly');
		img_name = $(this).attr('data-img-name');
		$('#selected-orings').html('<img class="ml-1 mr-2" src="https://s3.amazonaws.com/bodyartforms-products/' + img_name + '"   style="width:40px"><span class="mr-3">Qty: 4</span>' + friendly_title);
		$('#orings-dropdown-text').hide();
		});

		$('.remove_orings').click(function(){
		$('.free_orings').css('backgroundColor','#F5A9A9');
		$('.free_orings').fadeOut( "slow" );
		Cookies.set('orings', 'no', { expires: 30});
		});

	// Remove item from cart
	$(document).on('click', '.action-remove', function() {
		// remove stock notice if it's displaying
		$('.stock_notification').hide();
		
		row_cart_detailid = $(this).attr("data-detailid");
		var row_cart_specs = $(this).attr("data-specs");
		var new_qty = $("input[name=qty_change_id_" + row_cart_specs + "]").val();
		var orig_qty = $("input[name=qty_change_id_" + row_cart_detailid + "]").attr("data-orig_qty");
				
		var_cart_count = $('.cart-count').html();
		
	//	console.log(parseInt(var_mobile_cart_count) - parseInt(new_qty));
	//	console.log(parseInt(var_cart_count));
	//	console.log(parseInt(new_qty));
	
		
		$('.cart-count').html(parseInt(var_cart_count) - parseInt(new_qty));
		
		$.ajax({
			type: "post",
			url: "cart/ajax_cart_remove_item.asp",
			data: {cart_id: row_cart_detailid},
			success: function(){
				$('.detailid_' + row_cart_detailid).css('backgroundColor','#F5A9A9');
				$('.detailid_' + row_cart_detailid).fadeOut( "slow" );
				calcAllTotals();
			}	
		});
		
	});	


	// Save item for later
	$('.action-save-later').click(function(){
		row_cart_detailid = $(this).attr("data-detailid");
		var orig_qty = $("input[name=qty_change_id_" + row_cart_detailid + "]").attr("data-orig_qty");	
		var_cart_count = $('.cart-count').html();
		$('.cart-count').html(parseInt(var_cart_count) - parseInt(orig_qty));
		$.ajax({
			type: "post",
			url: "/cart/inc_cart_save_for_later.asp",
			data: {cart_id: row_cart_detailid, add_save: "yes"},
			success: function(){
				$("#saved-items").load("/cart/inc_cart_saved_items.asp"); // reload save for items 
				$('.detailid_' + row_cart_detailid).css('backgroundColor','#CEF6D8');
				$('.detailid_' + row_cart_detailid).fadeOut( "slow" );
				calcAllTotals();
			}	
		});
	});	


	// Add add-ons like salves, jewelry, etc to cart
	$('.add-cart-addon').click(function(){
		item_detailid = $(this).attr("data-detailid");

		$.ajax({
			type: "post",
			url: "/cart/ajax_cart_add_item.asp?qty=1&DetailID=" + item_detailid,
			data: {item_detailid: item_detailid},
			success: function(){
				$('#btn_' + item_detailid).html('<i class="fa fa-spinner fa-spin"></i> Adding...');
				location.reload();
			}	
		});
	});	

	// Update items in cart
		$(".qty_change").change(function() {
			// remove stock notice if it's displaying
			$('.stock_notification').hide();

			var qty_field_id = $(this).attr("id");
			qty_value = $('#' + qty_field_id).val();
			var orig_qty_cart = $('#' + qty_field_id).attr("data-orig_qty");
			var qty_cart_id = $('#' + qty_field_id).attr("id");
			row_cart_detailid = $('#' + qty_field_id).attr("data-detailid");
			var item_sale_price = $('#' + qty_field_id).attr("data-now_item_price");

			if (qty_value <= 0) {
				$('#' + qty_field_id).val('1');
				$('#' + qty_field_id).trigger('change');
			}
			

			var_cart_count = $('.cart-count').html();
	//	console.log("Cart count: " + var_cart_count + "     Orig qty: " + orig_qty_cart + "   New qty: " + qty_value);
	//	console.log(parseInt(orig_qty_cart) - parseInt(qty_value));
			if (qty_value > orig_qty_cart) {
			//	console.log("Increase");
				$('.cart-count').html(parseInt(var_cart_count) + (parseInt(qty_value) - parseInt(orig_qty_cart)));
			} else {
			//	console.log("Decrease");
				$('.cart-count').html(parseInt(var_cart_count) - (parseInt(orig_qty_cart) - parseInt(qty_value)));
			}
			
			function reWriteValues() {
				
				new_row_sale_price = qty_value * item_sale_price;				
				
			
				$(".success_id_" + row_cart_detailid).show();
					$(".success_id_" + row_cart_detailid).delay(3000).fadeOut();
					
					$('.qty_change_id_' + row_cart_detailid).attr('data-orig_qty',qty_value);
					
					$('.line_item_total_' + row_cart_detailid).attr('data-price',new_row_sale_price);
					
					// toFixed corrects format number decimal places
					$('.line_item_total_' + row_cart_detailid).html(new_row_sale_price.toFixed(2));
					
			} // end reWriteValues() function

			$.ajax({
			type: "post",
			dataType: "json",
			url: "cart/ajax_cart_update_item.asp?update=" + qty_cart_id + "&qty=" + qty_value + "&detailid=" + row_cart_detailid + "&orig_qty=" + orig_qty_cart + ""
			})
			//success: function(json){
			.done(function( json, msg ) {				

				// If qty in stock is less than what customer wanted re-write qty field				
				var requested_qty = qty_value
				var reset_qty = json.qty

				if (reset_qty != 0) {
					if (reset_qty != "out of stock") {
						qty_value = reset_qty
					} else { // change to 0 qty
						qty_value = 0
					}
				}
				
			// Pulls from json qty to reset stock ordered to
			function reWriteOverOrdered() {
				
			//	new_row_retail_price = reset_qty * item_retail_price;	
				new_row_sale_price = reset_qty * item_sale_price;				
			
			
				$(".success_id_" + row_cart_detailid).show();
					$(".success_id_" + row_cart_detailid).delay(3000).fadeOut();
					
					$('.qty_change_id_' + row_cart_detailid).attr('data-orig_qty',qty_value);		
					
					$('.line_item_total_' + row_cart_detailid).html('$' + new_row_sale_price.toFixed(2));
					
			} // end reWriteOverOrdered() function

			
			function runStockCheck() {
			//	console.log("Show stock notice   Json: " +reset_qty + "  Requested qty: " + requested_qty + " Qty value: " + qty_value);
				$('#stock_notice').show();
				
				// Check stock and load page
				$("#stock_notice").load( "cart/inc_stock_display_notice.asp?qty=" + qty_value + "&detailid=" + row_cart_detailid);			
			}
			
				// Unless it's over the request amount, all vales return 0 for the stock check

				if (reset_qty != 0) {
					
					// change to database qty value if stock is found
					if (reset_qty != "out of stock") {
						console.log("set to db value");
						$(".qty_change_id_" + row_cart_detailid).val(json.qty);
						reWriteOverOrdered();
						runStockCheck();
						
					} else { // change to 0 qty
						console.log("set to 0");
						$('.detailid_' + qty_field_id).css('backgroundColor','#F5A9A9');
						$('.detailid_' + qty_field_id).fadeOut( "slow" );
						
					//	new_row_retail_price = 0 * item_retail_price;
						reWriteOverOrdered();
						runStockCheck();
					}
				
				} else {
				//	console.log("DO NOT show stock notice   Json: " +reset_qty + "  Requested qty: " + requested_qty);
					
					$('#stock_notice').hide();
					reWriteValues();
				}
				
			calcAllTotals();
			});
		});	

	// Add autoclave service
	$('.btn-add-autoclave').click(function(){
		$('.btn-add-autoclave').html('Adding...');
		$.ajax({
			type: "post",
			url: "cart/ajax_cart_add_item.asp",
			data: {qty: "1", DetailID: "34356"},
			success: function(){
				location.reload();
			}	
		});
	});

	// Submit an active sale coupon when clicking the apply button -- shortcut method
	$('.coupon-shortcut').click(function(){	
		$('#frm-coupon').submit();
	});

	
	// Show preorder edit specs box for item clicked
	$('.edit-spec').click(function(){
		var id = $(this).attr("data-id");
		$(this).hide();
		$('.spec' + id + ', .update' + id + ', .cancel' + id).show();
	});	
	
	// Cancel specs update when clicked
	$('.cancel-spec').click(function(){
		var id = $(this).attr("data-id");
		$(this).hide();
		$('.spec' + id + ', .update' + id).hide();
		$('.edit' + id).show();
	});	

	// Update preorder specs after click
	$('.update-spec').click(function(){
		var id = $(this).attr("data-id");
		var specsvalue = $('.specvalue' + id).val();

		$('.update' + id + ', .spec' + id + ', .cancel' + id).hide();
		$('.specspin' + id).css("display","inline-block");
		$.ajax({
			type: "post",
			url: "cart/ajax-update-preorder-specs.asp",
			data: {cartid:id, specs:specsvalue}
			})
			.done(function(msg) {
				$('.specspin' + id).css("display","none");
				$('.edit' + id + ', .updateconfirm' + id).show();
				$('.spectext' + id).html(specsvalue);
				$('.specvalue' + id).val(specsvalue);
				$('.updateconfirm' + id).delay(2000).fadeOut("slow");
			})	
	});	
	
		
	// Automatically load up saved items on page load
	$("#saved-items").load("/cart/inc_cart_saved_items.asp");	

	// Page through saved for later items
	$(document).on("click", "#saved-items .page-link", function(event)
	{ 
		event.preventDefault();
		var saved_url = $(this).attr("href");
		$("#saved-items").load("/cart/inc_cart_saved_items.asp" + saved_url);
	});	

	// Load up items in modal to update cart item
	$(document).on("click", "#btn-edit-cart-item", function() {
		$('#btn-update-detail').html('Update item');
		var cartid = $(this).attr("data-cartid");
		var productid = $(this).attr("data-productid");
		var cart_qty = $('#' + cartid).val();
		
		$('#update-item-display-product').load("/cart/ajax-cart-update-item-display-product.asp", {productid: productid});
		$('#form-edit-cart-item').load("/products/ajax-details-dropdown-addtocart.asp", {productid: productid, cartid: cartid, cart_qty: cart_qty});
	});	

	// Change out selected drop down text for update item
	$(document).on("change", ".add-cart", function(event)
	{
		var selected_text = $(this).attr("dropdown-title");
		$('#selected-item').html(selected_text);
	});

	// Update / change out the item on the button press
	$(document).on("click", "#btn-update-detail", function(){
		$('#btn-update-detail').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');
		var_detailid = $('.add-cart:checked').val();
	//	console.log(var_detailid);

		$.ajax({
			type: "post",
			url: "cart/ajax-cart-update-item-detail.asp",
			data: $('#form-edit-cart-item').serialize() + "&detailid=" + var_detailid
			})
			.done(function(msg) {
				window.location = "/cart.asp?updateditem=yes";
			}) 	
	});
	
	// APPLY COUPON OR CERTIFICATE CODE
	$('#frm-coupon').submit(function(e) {
		var coupon_code = $('#coupon_code').val();
		if(coupon_code !=""){
			$('#coupon-applied').html('');
			$('#processing-message').show();
			$('#processing-message').html('<div class="alert alert-success mt-2"><i class="fa fa-spinner fa-2x fa-spin"></i> Please wait ... Applying coupon</div>');		
			$.ajax({
			  url: "/cart/ajax-cart-apply-coupon.asp",
			  type: 'POST',
			  data: {
				 coupon_code: coupon_code
			  }
			}).done(function( data, msg ) {
				$('#coupon-applied').html(data);
				calcAllTotals();
				$('#processing-message').html('');
			}).fail(function() {
				$('#processing-message').html('');
			});
	    }
		$('#coupon_code').val('');
	   e.preventDefault();
	}); 
	
	// Check if there is custom items in the cart
	$(document).on("click", "#btn-checkout, #btn-paypal", function(e){
		if(preOrderItem == "yes"){
			e.preventDefault();
			checkoutMethod = $(this).attr("id");
			$('#custom-order-warning-modal').modal('show');
			$('#custom-order-items').load("/cart/ajax-pre-order-item-display.asp");	
		}else{	
			// Proceed checkout process
		}
	});
	
	// Check if there is custom items in the cart
	$(document).on("click", "#btn-proceed-to-checkout", function(e){
		preOrderItem = "";
		$('#custom-order-warning-modal').modal('hide');
		if(checkoutMethod == "btn-checkout"){
			window.location = "/checkout.asp?type=card"
		}else if(checkoutMethod == "btn-paypal"){
			window.location = "/checkout.asp?type=paypal"
		}else if(checkoutMethod == "btn-googlepay"){
			onGooglePaymentButtonClicked();
		}else if(checkoutMethod == "btn-applepay"){		
			var appleButton = document.querySelector('apple-pay-button');
			appleButton.click();
		}
	});	
		