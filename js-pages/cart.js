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
	
	
	$(document).on('click', '#gaugeCardCheck', function() {
		if ($(this).is(':checked')){
			Cookies.set('gaugecard', 'yes', { expires: 30});
		}else{
			Cookies.set('gaugecard', 'no', { expires: 30});		
		}	
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

		$(document).on("change", 'input[class=freegift]', function(event) { 
			var_freegiftid = $(this).val();
			friendly_title = $(this).attr('data-friendly');
			slide_id = $(this).attr('data-slide-id');
			img_name = $(this).attr('data-img-name');
			var tier = $(this).attr('data-tier');

			$('#selected-gift' + tier).html(friendly_title);
			$('#gift' + tier + '-dropdown-text').hide();
			$('.selected-variation-text').hide();
			$('.variation-text').show(); 
			$('.slick-slide img').css("border", "none");
			if (friendly_title != "no free item"){
				Cookies.set('freegift' + tier + 'id', var_freegiftid, { expires: 30});
				Cookies.set('freegift' + tier + 'slideindex', $(this).attr('data-slide-index'), { expires: 30});
				Cookies.set('freegift' + tier + 'Title', friendly_title, { expires: 30});

				console.log("c:" + Cookies.get('freegift' + tier + 'id'));
				console.log("#slick-free-items-" + tier + ' .slick-slide[data-slide-index=' + $(this).attr('data-slide-index') + ']');
				
				var element = $("#slick-free-items-" + tier + ' .slick-slide[data-slide-index=' + $(this).attr('data-slide-index') + ']');
				element.find('.variation-text').hide();
				element.find('.selected-variation-text').html('<div class="small" style="color: #007bff">' + friendly_title + ' <i class="fa fa-check"></i></div>');
				$('#header-card-selected-item-' + tier).html(friendly_title + ' <i class="fa fa-check"></i>');
				element.find('.selected-variation-text').show();			
				element.find('img').css("border", "3px solid #1585ff");
				element.find('img').css("border-radius", "2px");
			}else{
				$('#header-card-selected-item-' + tier).html('');
				Cookies.set('freegift' + tier + 'id', "", { expires: 30});
				Cookies.set('freegift' + tier + 'slideindex', "", { expires: 30});
				Cookies.set('freegift' + tier + 'Title', "", { expires: 30});			
			}
			calcAllTotals();
			
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
			var item_anodization_subtotal = $('#' + qty_field_id).attr("data-anodization-subtotal");
			var var_anodizationBasePrice = $('#' + qty_field_id).attr("data-anodization-basePrice");
			console.log(var_anodizationBasePrice);
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
				
				new_row_sale_price = qty_value * item_sale_price;									new_row_anodization_price = qty_value * var_anodizationBasePrice;	

				$(".success_id_" + row_cart_detailid).show();
					$(".success_id_" + row_cart_detailid).delay(3000).fadeOut();
					
					$('.qty_change_id_' + row_cart_detailid).attr('data-orig_qty',qty_value);
					
					$('.line_item_total_' + row_cart_detailid).attr('data-price',new_row_sale_price);
					
					// toFixed corrects format number decimal places
					$('.line_item_total_' + row_cart_detailid).html(new_row_sale_price.toFixed(2));
					$('.anodization_line_total_' + row_cart_detailid).html(new_row_anodization_price.toFixed(2));
					
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
				new_row_anodization_price = reset_qty * var_anodizationBasePrice;	

				$(".success_id_" + row_cart_detailid).show();
					$(".success_id_" + row_cart_detailid).delay(3000).fadeOut();
					
					$('.qty_change_id_' + row_cart_detailid).attr('data-orig_qty',qty_value);		
					
					$('.line_item_total_' + row_cart_detailid).html('$' + new_row_sale_price.toFixed(2));
					$('.anodization_line_total_' + row_cart_detailid).html('$' + new_row_anodization_price.toFixed(2));
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
				$('.spectext' + id).html(sanitizeHTML(specsvalue));
				$('.specvalue' + id).val(specsvalue);
				$('.updateconfirm' + id).delay(2000).fadeOut("slow");
			})	
	});	

	function sanitizeHTML(text) {
	  var element = document.createElement('div');
	  element.innerText = text;
	  return element.innerHTML;
	}
	
		
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
		cart_qty = cart_qty.replace(/\D/g, ''); //Strip non-numeric chars
		
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
		
		
	// Load Free Items
	$(document).on("click", ".btn-free-items", function(event)
	{    
		$("#free-items").hide();	
		$("#loading-message").html('<div class="text-center alert" style="color: #525252; background-color: #f1f1f1; border-color: #dbdbdb;"><i class="fa fa-spinner fa-spin" style="font-size: 1.3em"></i> Loading free items . . .</div>');
		$("#loading-message").show();
		$.ajax({
		  url: "/cart/inc_freeitems_modal.asp",
		  type: 'POST',
		  data: {free_items_count: free_items_count}
		}).done(function( data, msg ) {
			$("#free-items").html(data);
			$("#free-items").show();
			$("#loading-message").hide();	
			$(".slick-free-items").slick({
					// Do not set slidesToShow property. It cannot detect parent's width since accordion is collapsed/hidden. Instead use variableWidth: true, and set a fixed width.
					// TODO: check this out in mobile phone, if view is ok. Setting a smaller width may be needed in mobile for slick items.
					/*
					slickGoTo: 1,
					focusOnSelect: true,
					centerMode: true,*/
					/*infinite: false,*/
					variableWidth: true,
					slidesToScroll: 3,					
					prevArrow: '<div class="slider-arrow-prev" style="height: 75%"><i class="fa fa-chevron-circle-left text-dark fa-2x pointer"></i></div>',
					nextArrow: '<div class="slider-arrow-next" style="height: 75%"><i class="fa fa-chevron-circle-right text-dark  fa-2x pointer"></i></div>'
			});	
			
			for (let i = 0; i < 7; i++) {
				let selectedSlickId = Cookies.get('freegift' + i + 'slideid');
				let selectedSlideIndex = Cookies.get('freegift' + i + 'slideindex');
				let friendly_title = Cookies.get('freegift' + i + 'Title');

				if(selectedSlideIndex != "" && selectedSlideIndex != undefined){		
					var element = $('#slick-free-items-' + i + ' .slick-slide[data-slide-index=' + (selectedSlideIndex) + ']');
					//var slickIndex = element.prev().attr("data-slick-index");
					//$("#slick-free-items-1").slick('slickGoTo', slickIndex, false);		
					element.find('.variation-text').hide();
					element.find('.selected-variation-text').html('<div class="small" style="color: #007bff">' + friendly_title + ' <i class="fa fa-check"></i></div>');
					$('#header-card-selected-item-' + i).html(friendly_title + ' <i class="fa fa-check"></i>');
					element.find('.selected-variation-text').show();			
					element.find('img').css("border", "3px solid #1585ff");
					element.find('img').css("border-radius", "2px");		
				}
			}
			
			}).fail(function() {
				$("#loading-message").hide();
				$("#free-items").show();
			});
	});	
		
	// Free Items Selection
	$(document).on("click", ".slick-slide", function(event)
	{   	
		if(!$(this).closest(".baf-carousel").hasClass("notavailable") && !$(this).closest(".baf-carousel").hasClass("free-stickers")){
			event.stopImmediatePropagation();
			event.preventDefault();
			
			var tier = $(this).attr("data-tier");
			var productid = $(this).attr("data-productid");
			var slideindex = $(this).attr("data-slide-index");
			
			$("#select-product-variation-" + tier).hide();
			$("#slick-free-items-" + tier).css("opacity", "1");
			
			$.ajax({
			  url: "/cart/inc_freeitems_select.asp",
			  type: 'POST',
			  data: {productid: productid, tier: tier, slideindex: slideindex}
			}).done(function( data, msg ) {
				$("#select-product-variation-" + tier).html(data);
				$("#select-product-variation-" + tier).show();
				$("#slick-free-items-" + tier).css("opacity", "0.5");
				// Check transition on production website and put a setTimeout if needed
			}).fail(function() {
				$("#select-product-variation-" + tier).html('');
			});	
		// If clicked on free stickers, don't load select menu
		}else if ($(this).closest(".baf-carousel").hasClass("free-stickers")){
			
			var productDetailId = $(this).attr("data-product-detail-id");
			friendly_title = $(this).attr('data-friendly');
			var tier = 2; // free stickers

			$('#selected-gift' + tier).html(friendly_title);
			$('#gift' + tier + '-dropdown-text').hide();
			$('.selected-variation-text').hide();
			$('.variation-text').show(); 
			$('.slick-slide img').css("border", "none");
			

			Cookies.set('freegift' + tier + 'id', productDetailId, { expires: 30});
			Cookies.set('freegift' + tier + 'slideindex', $(this).attr('data-slide-index'), { expires: 30});
			Cookies.set('freegift' + tier + 'Title', friendly_title, { expires: 30});
			
			var element = $("#slick-free-items-" + tier + ' .slick-slide[data-slide-index=' + $(this).attr('data-slide-index') + ']');
			element.find('.variation-text').hide();
			element.find('.selected-variation-text').html('<div class="small" style="color: #007bff">' + friendly_title + ' <i class="fa fa-check"></i></div>');
			$('#header-card-selected-item-' + tier).html(friendly_title + ' <i class="fa fa-check"></i>');
			element.find('.selected-variation-text').show();			
			element.find('img').css("border", "3px solid #1585ff");
			element.find('img').css("border-radius", "2px");
			calcAllTotals();		
		}
	});

	$(document).on("click", "body", function(event)
	{   	
		$(".select-product-variation").hide();
		if(!$(this).hasClass("slick-slide"))
			$(".slick-free-items").css("opacity", "1");
	});
	
	$(document).on("show.bs.collapse", ".collapse", function(e)	{
		let tier = $(this).attr("data-tier");
		let selectedSlickId = Cookies.get('freegift' + tier + 'slideid');
		let selectedSlideIndex = Cookies.get('freegift' + tier + 'slideindex');
		let friendly_title = Cookies.get('freegift' + tier + 'Title');

		if(selectedSlideIndex != "" && selectedSlideIndex != undefined){		
			var element = $("#slick-free-items-" + tier + ' .slick-slide[data-slide-index=' + (selectedSlideIndex) + ']');
			var slickIndex = element.prev().attr("data-slick-index");
			console.log("Go to slick Index:" + slickIndex);
			$("#slick-free-items-" + tier).slick('slickGoTo', slickIndex);		
			element.find('.variation-text').hide();
			element.find('.selected-variation-text').html('<div class="small" style="color: #007bff">' + friendly_title + ' <i class="fa fa-check"></i></div>');
			$('#header-card-selected-item-' + tier).html(friendly_title + ' <i class="fa fa-check"></i>');
			element.find('.selected-variation-text').show();			
			element.find('img').css("border", "3px solid #1585ff");
			element.find('img').css("border-radius", "2px");		
		}
	});
	
	function clearFreeItemsCookie(tier){
		for (let i = tier; i <= 7; i++) {
			Cookies.set('freegift' + i + 'id', "", { expires: 30});
			Cookies.set('freegift' + i + 'slideindex', "", { expires: 30});
			Cookies.set('freegift' + i + 'Title', "", { expires: 30});	
		}
	}

	