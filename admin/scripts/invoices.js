	// Clipboard 
	new Clipboard('.clipboard');

	// Initialize bootstrap popovers
	$(function () {
		$('[data-toggle="popover"]').popover()
	  })
	
	$('#hide_address').hide();
	

	function check_ups() {
		if ($('#shipping-type option:contains("UPS"):selected').length) {
			$('.usps-tracking').hide();
			$('.ups-tracking').show();
		} else {
			$('.usps-tracking').show();
			$('.ups-tracking').hide();
		}
	}
	check_ups();
	
	$('#shipping-type').change(function(){
		check_ups();
	});

	$('#show_address').click(function(){
		$('#div_shipping_address').show();
		$('#show_address').hide();
		$('#hide_address').show();
	});
	$('#hide_address').click(function(){
		$('#div_shipping_address').hide();
		$('#hide_address').hide();
		$('#show_address').show();
	});	
	
	// Expand all link click
	$(".expand").click(function(){	
		$(".show-less").toggleClass("details-border");
		$(".expanded-details").toggle('slide');
		var oldText = $(this).text();
		var newText = $(this).data('text');
		$(this).text(newText).data('text',oldText);
	}); // end expand all click
	
	function expand_one() {
		// Expand ONE click
		$(".expand-one").click(function(){	
			var id = $(this).attr("data-id");
			$('.expanded-details:not(.' + id + ')').hide();
			$("." + id).toggle('slide');
			$('tr').removeClass('table-info')
			if ($(".detail-main-" + id).hasClass('table-info')) {
				// Do nothing.. no classes need to be changed
			} else {
				$(".detail-main-" + id).toggleClass('table-info');
				$("." + id).toggleClass('table-info');
			}
			
		}); // end expand ONE click
	}
	
	expand_one();

	// Change submit backorder modal
	$(document).on("click", "#btn-update-bo-modal", function(event){
		var orderdetailid = $(this).attr("data-itemid");
		var qty = $(this).attr("data-qty");
		var title = $(this).attr("data-title");
		$("#frm-submit-backorder, #btn-submit-bo").show();
		$('#new-bo-message').html('');
		$('#btn-submit-bo').attr("data-itemid", orderdetailid);
		$("#radio").prop("checked", true);
		$('#bo-qty').prop('selectedIndex',0);
		$('#bo-qty-instock').html(qty);
		$('#bo-item-title').html(title);
	
	}); // End submit new backorder
	
	// Expand backorder
	$(document).on("click", ".process-bo", function(event){
		var item = $(this).attr("data-id");
		var qty = $(this).attr("data-qty");
		var detailid = $(this).attr("data-detailid");
		var price = $(this).attr("data-price");
		var origprice = $(this).attr("data-origprice");
		var qty_instock = $(this).attr("data-qty_instock");
		var invoice = $('#main-id').val();
		var total = $('#invoice-total').html();
		var card_number = $('#card_number').val();
		
		$('.load-bo').load('invoices/ajax-backorder-load.asp', {item: item, invoice: invoice, qty: qty, detailid: detailid, price: price, origprice: origprice, total: total, card_number: card_number, qty_instock: qty_instock});
	}); // end change submit backorder modal

	// Submit a new backorder
	$(document).on("click", "#btn-submit-bo", function(event){
		var orderdetailid = $(this).attr("data-itemid");
		var bo_reason = $("input:radio[name='BOReason']:checked").val();
		var bo_qty = $("#bo-qty").val();
		
		$("#frm-submit-backorder").hide();
		$('#new-bo-message').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');
		$('#btn-submit-bo').hide();

		$.ajax({
		method: "POST",
		url: "invoices/ajax-backorder-submit.asp",
		data: {orderdetailid:orderdetailid, bo_reason:bo_reason, bo_qty:bo_qty}
		})
		.done(function(msg ) {
			$("#new-bo-message").html('<span class="alert alert-success">Backorder processed</span>').show();
			$('.bo_blue_' + orderdetailid).hide();
			$('.bo_orange_' + orderdetailid).show();
		})
		.fail(function(msg) {
			alert('FAILED');
			$("#new-bo-message").hide();
			$('#btn-submit-bo').show();
			$("#frm-submit-backorder").delay(1000).fadeIn(1000);
		});
	}); // End submit new backorder
	
	// Expand backorder
	$(document).on("click", ".process-bo", function(event){
		var item = $(this).attr("data-id");
		var qty = $(this).attr("data-qty");
		var detailid = $(this).attr("data-detailid");
		var price = $(this).attr("data-price");
		var origprice = $(this).attr("data-origprice");
		var qty_instock = $(this).attr("data-qty_instock");
		var invoice = $('#main-id').val();
		var total = $('#invoice-total').html();
		var card_number = $('#card_number').val();
		
		$('.load-bo').load('invoices/ajax-backorder-load.asp', {item: item, invoice: invoice, qty: qty, detailid: detailid, price: price, origprice: origprice, total: total, card_number: card_number, qty_instock: qty_instock});
	}); // end Expand backorder

	
	// Backorder processing
	$(document).on("click", ".btn_bo", function(event){
		var item = $(this).attr("data-item");
		var agenda = $(this).attr("data-agenda");
		var qty = $('#qty_' + item).val();
		var productid = $('#detailid_' + item).val();
		var detailid = $('#detailid_' + item).val();
		var exchange_origitem = $(this).attr("data-origitem");
		var exchange_detailid = $('#bo-exchange-detailid').val();
		var exchange_productid = $('#bo-exchange-product').val();
		var exchange_qty = $('#bo-exchange-qty').val();
		var exchange_price_diff = $('#bo-exchange-price-diff').html();
		var exchange_agenda = $(this).attr("data-exchange_agenda");
		var price = $('#price_' + item).val();
		var origprice = $('#origprice_' + item).val();
		var invoice = $('#main-id').val();
		var total = $('#invoice-total').html();
		var card_number = $('#card_number').val();
		
	//	console.log(exchange_price_diff);
		$('.bo-message').html('<i class="fa fa-spinner fa-2x fa-spin"></i>').show();
		$(".backorders").hide();
		
		$.ajax({
		method: "POST",
		dataType: "json",
		url: "invoices/ajax-backorder-process.asp",
		data: {item: item, agenda: agenda, invoice: invoice, qty: qty, detailid: detailid, price: price, total: total, origprice: origprice, card_number: card_number, exchange_detailid: exchange_detailid, exchange_productid: exchange_productid, exchange_qty: exchange_qty, exchange_origitem: exchange_origitem, exchange_price_diff: exchange_price_diff, exchange_agenda: exchange_agenda}
		})
		.done(function( json, msg ) {
			$('.bo-message').html('<div class="alert alert-success">' + json.status + '</div><br/>').show();
			$('.bo_blue_' + item).show();
			$('.bo_orange_' + item).hide();
			$(".backorders").delay(3000).fadeIn(1000);
		//	$(".bo-message").delay(10000).fadeOut(1000);
		})
		.fail(function(msg) {
			alert('FAILED');
			$(".bo-message").hide();
			$(".backorders").delay(1000).fadeIn(1000);
		});
	}); // End Backorder processing
	
	// Exchange BO, search product by #
	$(document).on("change", "#bo-exchange-product", function(event){
		var productid = $('#bo-exchange-product').val();
		$('#exchange-results').load('invoices/ajax_load_search_results.asp?productid=' + productid);
	});	//  END Exchange BO, search product by # ---------------------------------------
	
	// BO Calc price differences for exchange
	function calcExchangeDiff() {
		var price_diff = (($('#bo-exchange-price').val() * $('#bo-exchange-qty').val()) - $('#bo-exchange-origprice').html()).toFixed(2);
		
		if (price_diff < 0) {
			$('#price-diff-label').html('Refund due: ');
			$('#exchange-agendas').html('<input type="radio" name="exchange-refund" value="cardrefund" class="exchange-agenda"> Refund &nbsp;&nbsp;&nbsp;<input type="radio" name="exchange-refund" value="storecredit" class="exchange-agenda"> Store credit<br/><br/>');
			$('#exchange-agendas').show();
			$('#btn-exchange').hide();
			$('#btn-exchange').attr('data-exchange_agenda', '');
		} 
		if (price_diff > 0) {
			$('#price-diff-label').html('Amount due: ');
			$('#exchange-agendas').html('');
			$('#exchange-agendas, #btn-exchange').show();
			$('#btn-exchange').attr('data-exchange_agenda', 'amount-owed');
		} 
		if (price_diff == 0) {
			$('#price-diff-label').html('Price difference: ');
			$('#btn-exchange').show();
			$('#exchange-agendas').hide();
			$('#btn-exchange').attr('data-exchange_agenda', 'equal-exchange');
		} 
		
		$('#bo-exchange-price-diff').html(price_diff);
		if (price_diff < 0) {
			$('#bo-exchange-price-diff').html(price_diff * -1); // set to positive amount just for display purposes
		}
		
		$('#btn-exchange').attr('data-price', $('#bo-exchange-price-diff').html());
	}
	
	// Write exchange agenda to Exchange button attribute
	$(document).on("change", '.exchange-agenda', function() {
		$('#btn-exchange').attr('data-exchange_agenda', $(this).val());
		$('#btn-exchange').show();

		console.log('radio value ' + $(this).val());
	});
	
	// BO Exchange select item to exchange for
	$(document).on("click", '.bo-exchange-item', function(event) { 
		item_price = $(this).attr("data-add_price")
		discount_rate = $('#bo-discount-rate').val();
		$('#bo-exchange-form').show();
		$('#bo-exchange-detailid').val($(this).attr("data-add_detail"));
		$('#bo-exchange-itemname').html($(this).html() + "<br/>");
		$('#btn-exchange').attr('data-detailid', $('#bo-exchange-detailid').val());
		
		
		// if discount found, calculate savings
		if ($('#bo-discount-rate').val() > 0) {
			$('#bo-exchange-price').val((item_price - ((discount_rate / 100) * item_price)).toFixed(2));
		} else { // if not discount, display regular price
			$('#bo-exchange-price').val($(this).attr("data-add_price"));
		}
		
		// Display price difference (if any)
		calcExchangeDiff();
	});
	
	// BO calc price diff if price is changed
	$(document).on("change", '#bo-exchange-price, #bo-exchange-qty', function(event) { 
		// Display price difference (if any)
		calcExchangeDiff();
	});
	



	
	// Delete item
	$('.delete_item').click(function(){
		var detailid = $(this).attr("data-delete_id");
		var invoiceid = $('#main-id').val();
		var item_price = $(this).attr("data-price");
		var item_origprice = $(this).attr("data-origprice");
		var qty = $(this).attr("data-qty");
		$(this).html('<i class="fa fa-spinner fa-spin"></i>');
		
		$.ajax({
		method: "POST",
		url: "invoices/ajax_delete_item.asp",
		data: {detailid: detailid, invoiceid: invoiceid, item_price: item_price, item_origprice: item_origprice, qty: qty}
		})
		.done(function( msg ) {
			$('#tbody-' + detailid).hide().fadeOut(5000);
			window.location.replace("?notice-type=add&id=" + invoiceid);
		})
		.fail(function(msg) {
			alert('Delete FAILED');
		});
	});

	// Update inventory
	$('.update_inventory').click(function(){
		var type = $(this).attr("data-type");
		var invoiceid = $('#main-id').val();
		
		$.ajax({
		method: "POST",
		url: "invoices/ajax_update_inventory.asp",
		data: {type: type, invoiceid: invoiceid}
		})
		.done(function( msg ) {
			$('#confirm_inv_updates').show().fadeOut(15000);
			load_notes();
		})
		.fail(function(msg) {
			alert('Delete FAILED');
		});
	});	
	
	
	// Change active / inactive drop down select colors
	$(".status").change(function(){
		if ($(this).val() == '1') {
			$(this).addClass('alert-success');
			$(this).removeClass('alert-danger');			
		} else {
			$(this).addClass('alert-danger');
			$(this).removeClass('alert-success');
		}
	}); // end active selector colors	
	
	// Load tracking information into page & display
	$('.usps_tracking').hover(
	  function() {
		var url = $(this).attr("data-url");
		var tracking_num = $('#usps-tracking').val();
		$('#tracking_display').load(url + tracking_num);
	  }, function() {
	  }
	);
	
	// Maintain tracking display status after click
	$('.usps_tracking').click(function() {
		$('#tracking_display').toggle();
		$('#tracking_arrow_up, #tracking_arrow_down').toggle()
	});

	// Update customer credit
	$('#customer_credit').change(function(){
		var custid = $(this).attr("data-custid");
		var amount = $(this).val();
		
		$.ajax({
		method: "POST",
		url: "customers/ajax_edit_customer_credit.asp",
		data: {custid: custid, amount: amount}
		})
		.done(function( msg ) {
			$('#confirm_credit_update').show().fadeOut(5000);
		})
		.fail(function(msg) {
			alert('Update FAILED');
		});
	});

	// Auto load notes and create function
	function load_notes() {
		var invoiceid = $('#main-id').val();
		$('#display_notes').load('invoices/ajax_get_notes.asp?id=' + invoiceid);
	};
	
	load_notes();

	
	// Update PRIVATE NOTES
	$('#private_notes').change(function(){
		var invoiceid = $('#main-id').val();
		var note = $(this).val();
		
		$.ajax({
		method: "POST",
		url: "invoices/ajax_add_note.asp",
		data: {invoiceid: invoiceid, note: note}
		})
		.done(function( msg ) {
			$('#private_notes').val('');
			load_notes();
		})
		.fail(function(msg) {
			alert('Update FAILED');
		});
	});



localStorage["move_details"] = "" // set initial value to nothing
localStorage["copy_details"] = "" // set initial value to nothing
	
	// Copy button
	$(".copyid").click(function(){	
		console.log("copy item");	
		var id = $(this).attr("data-id");
		var qty = parseInt($(this).attr("data-qty"));
		var qty_instock = parseInt($(this).attr("data-qty_instock"));
		$('.move-copy-productid').show();
		$('#move-copy-text').html("Copy");		
		$('.move-copy-productid').removeClass("move-detail");
		$('.move-copy-productid').addClass("copy-detail");
		$("button[name=copy_" + id + "]").toggleClass('btn-success');
		$("button[name^=move_]").removeClass('btn-success');
		
		if (qty_instock < qty){
			alert("WARNING: We only have " + qty_instock + " in stock. Invoice has " + qty + " ordered. You can still copy the item, but please double check stock levels.")
		}
		
		if ($(".btn-success")[0]){} // if class if found do nothing 
		else { // if it's not found then hide span input box
			$('.move-copy-productid').hide();
		}
	
		// check for duplicates and dont' allow them
		if ($(this).hasClass('btn-success')) {
			// check for duplicates and dont' allow them
			if (localStorage.copy_details.indexOf(id) === -1) {
				localStorage["copy_details"] = localStorage.copy_details + id + ","
			}
		} else { // if it's inactive then remove the id from storage
					localStorage["copy_details"] = localStorage["copy_details"].replace(id + ',','');
		}	

	}); // Copy button
	
	// Move button
	$(".moveid").click(function(){
		console.log("move item");		
		var id = $(this).attr("data-id");
		$('.move-copy-productid').show();
		$('#move-copy-text').html("Move");
		$('.move-copy-productid').removeClass("copy-detail");
		$('.move-copy-productid').addClass("move-detail");
		$("button[name=move_" + id + "]").toggleClass('btn-success');		
		$("button[name^=copy_]").removeClass('btn-success');
		
		if ($(".btn-success")[0]){} // if class if found do nothing 		
		else { // if it's not found then hide span input box
			$('.move-copy-productid').hide();
		}
		
		if ($(this).hasClass('btn-success')) {
			// check for duplicates and dont' allow them
			if (localStorage.move_details.indexOf(id) === -1) {
				localStorage["move_details"] = localStorage.move_details + id + ","
			}
		} else { // if it's inactive then remove the id from storage
					localStorage["move_details"] = localStorage["move_details"].replace(id + ',','');
		}	
	}); // Move button

	// ------------ When product # is inputted for copy/move then load ajax and redirect to new page
	$(".move-copy-productid input[name=toggle-productid]").change(function(){
		var productid = $(".move-copy-productid input[name=toggle-productid]").val();
		var invoiceid = $('#main-id').val();
		
		
		if ($('#ReturnMailer').is(":checked")) {
			console.log("Send return mailer");
			var return_mailer = "yes"
		} else {
			console.log("No return mailer");
			var return_mailer = "no"
		}
		
		if ($('#reship-returned').is(":checked")) {
			console.log("Ship returned order");
			var reship_returned = "yes"
		} else {
			console.log("Do not ship returned order");
			var reship_returned = "no"
		}
		
		if ($('.move-copy-productid').hasClass('copy-detail')) {
			var toggle_type = "copy" 
			var details = localStorage["copy_details"]
		} else {
			var toggle_type = "move"
			var details = localStorage["move_details"]
		}	
		//	console.log(productid + ", " + toggle_type + ", " + details);
			
		$.ajax({
		method: "POST",
		dataType: "json",
		url: "invoices/ajax_duplicate_move_items.asp",
		data: {move_to_id: productid, details: details, toggle_type: toggle_type, invoiceid: invoiceid, return_mailer: return_mailer, reship_returned: reship_returned}
		})
		
		.done(function( json, msg ) {
			localStorage.removeItem("move_details");
			localStorage.removeItem("copy_details");
		
			if (toggle_type === "copy") {
				window.location.replace("?notice-type=deduct&id=" + json.invoiceid);
			} else {
				window.location.replace("?id=" + json.invoiceid);
			}
		})
		.fail(function(msg) {
			alert(toggle_type + " FAILED");
		});
	
	}); // END transferring copied & moved items ---------------------------------------	
	
	// Create new order button
	$("#create_new_order").click(function(){
		$('.move-copy-productid').show();
	});	
	
	$(".close").click(function(){
		var invoiceid = $('#main-id').val();
		$(this).closest("div").hide();
		window.location.replace("?id=" + invoiceid);
	});	


	// ADD item to order
	// ------------ When product # is inputted for copy/move then load ajax and redirect to new page
	$("#btn_add_item").click(function(){
		$("#btn_add_item").prop('disabled', true);
		$("#btn_add_item").html('<i class="fa fa-spinner fa-spin"></i>');

		var invoiceid = $('#main-id').val();
		var add_productid = $('#frm_productid').val();
		var add_detailid = $('#frm_detailid').val();
		var add_qty = $('#frm_qty').val();
		var add_price = $('#frm_price').val();

		$.ajax({
		method: "POST",
		url: "invoices/ajax_add_item.asp",
		data: {invoiceid: invoiceid, add_detailid: add_detailid, add_productid: add_productid, add_qty: add_qty, add_price: add_price}
		})
		
		.done(function( msg ) {
				window.location.replace("?notice-type=deduct&id=" + invoiceid);
		})
		.fail(function(msg) {
			alert(toggle_type + " FAILED");
		});
	
	}); // END add item to order ---------------------------------------		
	
	// BEGIN Pull up product search results
	$("#frm_search_product").change(function(){
		var productid = $('#frm_search_product').val();
		$('#show_frm_add').show();
		$('#search_results').load('invoices/ajax_load_search_results.asp?productid=' + productid);
	});	//  END Pull up product search results ---------------------------------------
	
	// Select search detail and pre-fill item price box and detail ID
	$(document).on("click", '.select_item_to_add', function(event) { 
		$('#frm_price').val($(this).attr("data-add_price"));
		$('#frm_detailid').val($(this).attr("data-add_detail"));
	});
		
	// Insert invoice total into all backorder links
	var invoice_total = $('#invoice-total').html();
	$("a.bo-ship").attr("href", function(i, href) {
		return href + '&total=' + invoice_total;
	});
	
	// Clicking icons writes notes into public notes field for scanner purposes
	$(document).on("click", '.insert-notes', function(event) { 
		var_insert_text = $(this).attr("data-text");
		$('#public_notes').val($('#public_notes').val() + '<br/>' + var_insert_text);
		$("#public_notes").trigger("change");
	});


	// Use terminal
	$('#frm-terminal').submit(function (event) {
		event.preventDefault(event); // Do not reload page

		$("#frm-terminal").hide();
		$('#message-frm-terminal').html('<i class="fa fa-spinner fa-2x fa-spin"></i>').show();
		terminal_pay_method = $('#pay-method').val();

		if (terminal_pay_method === 'Afterpay') {
			terminal_url = 'ajax-afterpay-terminal.asp';
		} else {
			terminal_url = 'ajax-authnet-transact.asp';
		}

		$.ajax({
		method: "post",
		dataType: "json",
		url: "terminal/" + terminal_url,
		data: $("#frm-terminal").serialize()
		})
		.done(function(json, msg) {
			if(json.status == "success") {
				$('#message-frm-terminal').html('<div class="alert alert-success">' + json.reason + '</div>').show();
				$('#message-frm-terminal').delay(7000).fadeOut('slow');
				$('#frm-terminal').delay(8000).fadeIn('slow');
			} else {
				$('#message-frm-terminal').html('<div class="alert alert-danger">' + json.reason + '</div>').show();
				$('#message-frm-terminal').delay(7000).fadeOut('slow');
				$('#frm-terminal').delay(8000).fadeIn('slow');
			}			
		})
		.fail(function(json, msg) {
			$('.message-cim-form').html('<div class="alert alert-danger">Website error.</div>').show();
		})
		return false;
	});  // Use terminal	


	// Set a temp customer id session
	$(document).on("click", '#temp-account', function(event) { 
		var_custid = $(this).attr("data-custid");
		
		$.ajax({
		method: "post",
		url: "/accounts/ajax_set_temp_custid_session.asp",
		data: {custid: var_custid}
		})
		.done(function(msg) {
			window.open('/account.asp', '_blank');		
		})
	});	  

	// Returns agenda
	$(document).on("click", '.return-agenda', function(event) { 
		var agenda = $(this).attr("data-agenda");
		var target1 = $(this).attr("data-show");
		var target2 = $(this).attr("data-show2");
		$('#submit-return').attr("data-agenda", agenda);
		$('.return-hide').hide();
		$('.' + target1 + ', .' + target2).show();
			$('.return-qty').prop('disabled', true);
			$('.return-check').removeClass('fa-plus-circle btn-success');
			$('.return-check').addClass('fa-times-circle btn-danger');
	});	

	// Hide show checkmark for refund shipping toggle
	$(document).on("click", '#btn-refund-shipping', function(event) { 
		$(this).toggleClass("btn-secondary btn-primary");
		$('#icon-toggle-shipping-refund').toggle();	
	});	

	// Select all / none returns
	$(document).on("click", '#btn-return-selectall', function(event) { 

		if ($(this).hasClass('fa-plus-circle')) {
			$(this).removeClass('fa-plus-circle');
			$(this).addClass('fa-times-circle');
			$('.return-qty').prop('disabled', false);
			$('.return-check').removeClass('fa-times-circle btn-danger');
			$('.return-check').addClass('fa-plus-circle btn-success');
		} else {
			$(this).removeClass('fa-times-circle');
			$(this).addClass('fa-plus-circle');
			$('.return-qty').prop('disabled', true);
			$('.return-check').removeClass('fa-plus-circle btn-success');
			$('.return-check').addClass('fa-times-circle btn-danger');
		}
	});	
	
	// Enable / Disable qty field for return
	$(document).on("click", '.return-check', function(event) { 
		var itemid = $(this).attr("data-id");
		if ($('#return-id-' + itemid).prop('disabled')) {
			$('#return-id-' + itemid).prop('disabled', false);
			$(this).removeClass('fa-times-circle btn-danger');
			$(this).addClass('fa-plus-circle btn-success');
		} else {
			$('#return-id-' + itemid).prop('disabled', true);
			$(this).removeClass('fa-plus-circle btn-success');
			$(this).addClass('fa-times-circle btn-danger');
		}
	});	
	

	// Returns - undeliverable email
	$(document).on("change", '#undeliverable-reason', function(event) { 
		$('#submit-return').show();
		if ($('#undeliverable-reason').val() === 'Other') {
			$('#group-undeliverable-other').show();
		} else {
			$('#group-undeliverable-other').hide();
		}
	});

	// Close returns modal
	$(document).on("click", '#btn-returns-close', function(event) { 
		$('.return-hide').hide();
		$('#message-returns').html('');
	});	
	
	// Calculate return amount
	$(document).on("click", '#btn-return-calculate', function(event) { 
		$.ajax({
			method: "post",
			url: "invoices/ajax-returns-calculate.asp",
			dataType: "json",
			data: $("#form-returned-items-selection").serialize()
			})
			.done(function(json, msg) {
				$('#message-returns').html('<div class="alert alert-success" id="returns-copy-msg">' + 
				'Subtotal ( Less coupon discount of $' + json.coupon_discount.toFixed(2) + ' and - $' + json.preorder_restock_fee.toFixed(2) + ' restock fee and - $' + json.db_free_use_now_credits.toFixed(2) + ' free use now credits )  $' + json.subtotal.toFixed(2) + 
				'<br/>New total ( Subtotal $' + json.subtotal.toFixed(2) + 
				' + shipping $' + json.shipping_rate.toFixed(2) + 
				' + additional amt $' + json.additional_amount.toFixed(2) +
				' + tax $' + json.sales_tax.toFixed(2) + 
				' ) $' + json.subtotal_plus_shipping_and_salestax.toFixed(2) + 
				'<br/>Auth.net original charge $' + json.authnet_settleAmount.toFixed(2) +
				'<br/>Gift cert refund due ( From code ' + json.gift_cert_code + 
				' | Invoice ' + json.gift_cert_invoice +' | $' + json.db_gift_cert.toFixed(2) + 
				' available ) <strong>$' + json.gift_cert_refund_due.toFixed(2) + 
				'</strong><br/>Store credit refund due ( $'+ json.db_store_credit.toFixed(2) +' available ) <strong>$' + json.store_credit_refund_due.toFixed(2) + 
				'</strong><br/>Credit card refund due: <strong>$' + json.cc_refund_due.toFixed(2) +
				'</strong><br/></div>');
				$('#submit-return').show();
				$('#returns-ccrefund').val(json.cc_refund_due.toFixed(2));
				$('#returns-storecredit-due').val(json.store_credit_refund_due.toFixed(2));
				$('#returns-giftcert-due').val(json.gift_cert_refund_due.toFixed(2));
				$('#returns-sales-tax').val(json.sales_tax.toFixed(2));
				$('#returns-calculation').val($('#returns-copy-msg').html());
				$('#returns-total').html(json.cc_refund_due + json.store_credit_refund_due + json.gift_cert_refund_due );
			})
			.fail(function(json, msg) {
				$('#message-returns').html('<div class="alert alert-danger">ERROR</div>');
			})
	});	

	
	// Submit and process returns
	$(document).on("click", '#submit-return', function(event) { 
		var agenda = $(this).attr("data-agenda");
		var invoice = $(this).attr("data-invoiceid");
		var cc_refund_due = $(this).attr("data-ccrefund");
		var storecredit_refund_due = $(this).attr("data-storecredit-due");
		var giftcert_refund_due = $(this).attr("data-giftcert-due");
        
        if (agenda === 'undelivered') {
			var reason = $('#undeliverable-reason').val();
			
			if (reason === 'Other') {
				var reason = $('#undeliverable-reason-other').val();
			}
			
            $.ajax({
                method: "post",
                url: "invoices/ajax-returns-undelivered-package.asp",
                data: {invoice: invoice, reason:reason}
                })
                .done(function(msg) {
                    $('#message-returns').html('<div class="alert alert-success">Email sent</div>');
                    $('#submit-return').hide();
					$("#order-status option[value='PACKAGE CAME BACK']").prop('selected', true);
					load_notes();
                })
                .fail(function(msg) {
                    $('#message-returns').html('<div class="alert alert-danger">Email failed</div>');
                })
		}
		
		        
        if (agenda === 'return-items') {
			
            $.ajax({
				method: "post",
				dataType: "json",
                url: "invoices/ajax-returns-process.asp",
				data: $("#form-returned-items-selection").serialize() + "&invoice=" + invoice + "&storecredit_refund_due=" + storecredit_refund_due + "&giftcert_refund_due=" + giftcert_refund_due + "&cc_refund_due=" + cc_refund_due
                })
                .done(function(json,msg) {
					if (json.status == 'success') {
						$('#message-returns').html('<div class="alert alert-success">SUCCESS</div>');
						$('#submit-return').hide();
						if (json.var_active_status != '') {
							$('#message-returns').append('<div class="alert alert-warning h5">INACTIVE PRODUCT(S) -- NEED NEW LABEL CREATED:<br>' + json.var_active_status + '</div>');
						}
					} else {
						$('#message-returns').html('<div class="alert alert-danger">CREDIT CARD / PAYPAL REFUND FAILED' + json.status + '</div>');
						$('#submit-return').hide();
						if (json.status_store_credit == 'success') {
							$('#message-returns').append('<div class="alert alert-success h6">STORE CREDIT HAS BEEN SUCCESSFULLY APPLIED</div>');
						}
						if (json.status_gift_cert == 'success') {
							$('#message-returns').append('<div class="alert alert-success h6">GIFT CERTIFICATE CREDIT HAS BEEN SUCCESSFULLY APPLIED</div>');
						}
					}
                })
                .fail(function(json,msg) {
					$('#message-returns').html('<div class="alert alert-danger">Failed</div>');
					$('#submit-return').hide();
                })
        }
	});	

	// Request shipping label
	$("#request-label").click(function(e) {
        var url = $(this).attr("data-url");
		var shipper = $(this).attr("data-shipper");
		$(this).html("Processing...");
		$('#label-message').html('');

		$.ajax({   
            method: "post",
		    url: url,
		    dataType: "json"
		})
            .done(function(json, msg) {

                if(json.status == 'error') {
					$("#request-label").html("CODE ERROR");
					$('#label-message').html('<div class="alert alert-danger p-2">' + json.message + '</div>');
                } else {
					$("#request-label").removeClass('d-inline-block');
					$("#request-label").hide();
					$("#reprint-label").html('Print ' + shipper + ' label');
					$("#reprint-label").show();
					$('#label-message').html('<div class="alert alert-success p-2">' + json.message + '</div>');
				}
            })
            .fail(function(json, msg) {
                $(this).html("CODE ERROR");
        })
		});
		

	// Change reship order modal
	$(document).on("click", ".btn-update-reship-modal", function(event){
		var invoiceid = $(this).attr("data-invoiceid");
		$('.btn-reship-items').attr("data-invoiceid", invoiceid);

		$('#load-reship-items').load("invoices-cs-modules/ajax-item-problems.asp", {invoiceid: invoiceid})
	
	}); // End change reship order modal

	// Approve or deny reships
	$(document).on("click", ".btn-reship-items", function(event){
		var invoiceid = $(this).attr("data-invoiceid");
		var agenda = $(this).attr("data-agenda");

		if(agenda==='approve') {
			$('#btn-reship-approve').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');
			$('#btn-reship-deny').hide();
		} else {
			$('#btn-reship-deny').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');
			$('#btn-reship-approve').hide();
		}
		$('#load-reship-items').hide();

		$.ajax({
		method: "post",
		url: "invoices-cs-modules/ajax-reship-items.asp",
		data: {invoiceid: invoiceid, agenda:agenda}
		})
		.done(function(msg ) {
			if(agenda==='approve') {
				$('#message-reship-status').html('<div class="alert alert-success h6">ITEMS HAVE BEEN PROCESSED</div>');
			} else {
				$('#message-reship-status').html('<div class="alert alert-success h6">WINDOW CAN BE CLOSED</div>');
			}
			
			$('.btn-reship-items').hide();
		})
		.fail(function(msg) {
			$('#message-reship-status').html('<div class="alert alert-danger h5">ERROR</div>');
			$('.btn-reship-items').hide();
		});
	
	}); // End Approve or deny reships

	// Resend shipment email
	$(document).on("click", '#btn-send-shipment-email', function() { 
		$(this).hide();
		$('#msg-send-shipment-email').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');

		var invoiceid = $(this).attr("data-invoiceid");

		$.ajax({
			method: "post",
			url: "update_multiple_records.asp",
			data: {checkbox: invoiceid}
			})
			.done(function(msg ) {
				$('#msg-send-shipment-email').html('<span class="alert alert-success p-2">E-mail sent</span>').delay(5000).fadeOut(1000);;
			})
			.fail(function(msg) {
				$('#msg-send-shipment-email').html('<span class="alert alert-danger p-2">E-mail failed</span>');

			});
	});	

	// BEGIN Duplicate invoice
	$(document).on("click", "#duplicate_order", function(){
		var invoiceid = $(this).attr("data-invoiceid");
		var email = $(this).attr("data-email");
		$('#msg-duplicate-order').html('<i class="fa fa-spinner fa-spin"></i>');

		$.ajax({
		dataType: "json",
		method: "post",
		url: "invoices/ajax-duplicate-invoice.asp",
		data: {invoiceid: invoiceid, email: email}
		})
		.done(function(json, msg ) {
			$('#msg-duplicate-order').html('<i class="fa fa-spinner fa-spin mr-3"></i>Transferring to new invoice...');
			window.location.replace("?id=" + json.new_invoiceid);
		})
		.fail(function(json, msg) {
			$('#msg-duplicate-order').html(' -- ERROR');
		});
	}); // END Duplicate invoice

	new Clipboard('#copy-order'); // Copies order to clipboard

	// Move copy order button to top of page.
	$('#copy-order').appendTo('#holder-copy-order')

// LEAVE AT BOTTOM OF PAGE	
	// Generic show target 
	$(".btn-show-target").click(function(){	
		var target = $(this).attr("data-show");
		$('.' + target).show();
	});


