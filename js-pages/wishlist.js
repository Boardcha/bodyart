// Keyword search
	$('#btn-submit-filters').click(function () {
		$.ajax({
		method: "post",
		url: "wishlist/ajax-wishlist-set-filter-sessions.asp",
		data: $(this).closest('form').serialize()
		})
		.done(function(msg) {
			window.location = "wishlist.asp?userID=" + $('#user-id').val();
		})
		.fail(function(msg) {
			console.log("failed");
		})
	});	
	
	
	// Filter form submit
	$('#wishlist-sort, #wishlist-jewelry, #wishlist-gauge, #wishlist-list, #wishlist-material, #wishlist-brand').change(function () {
		$.ajax({
		method: "post",
		url: "wishlist/ajax-wishlist-set-filter-sessions.asp",
		data: $(this).closest('form').serialize()
		})
		.done(function(msg) {
			window.location = "wishlist.asp?userID=" + $('#user-id').val();
		})
		.fail(function(msg) {
			console.log("failed");
		})
	});  // Filter form submit
	

	// Clear filters
	$('.clear-filters').click(function () {

		$.ajax({
		method: "post",
		url: "wishlist/ajax-wishlist-set-filter-sessions.asp",
		data: {clear:"yes"}
		})
		.done(function(msg) {
			window.location = "wishlist.asp?userID=" + $('#user-id').val();
		})
		.fail(function(msg) {
			console.log("failed");
		})
	});	// Clear filters
	
	// For mobile, detect if a select filter has a value and if it does then keep the filters expanded
		var val_jewelry = $( "#wishlist-jewelry" ).val();
		var val_gauge = $( "#wishlist-gauge" ).val();
		var val_brand = $( "#wishlist-brand" ).val();
		var val_material = $( "#wishlist-list" ).val();
		
		console.log("jewelry: " + val_jewelry + "gauge: " + val_gauge + "brand: " + val_brand + "material: " + val_material)
		
		if (val_jewelry != '' || val_gauge != '' || val_brand != '' || val_material != ''){
		
		
			$('#expand-filters').show();
			$('.gallery-filter-down, .gallery-filter-up').toggle();

		}
	
	// Add to cart
	$('.add-cart').click(function () {
		var wishlist_id = $(this).attr('data-id');
		var DetailID = $(this).attr('data-detailid');
		var comments = $('#comments-' + wishlist_id).html();
		var ProductID = $(this).attr('data-productid');
		var qty = $('#desired-' + wishlist_id).html();

		if (comments != '') {
			var customorder = "yes"
		}
		
		
		if (getCookie("cartCount") > 0) {
			var cart_count = $('.cart-count').html();
		} else {
			var cart_count = 0;
		}

		$(this).hide();
		$('.msg-add-cart-' + wishlist_id).html('<div class="p-1 my-2">Adding items ...</div>').show();

	//	console.log("desired " + qty + "  cart count " + cart_count);
		
		$.ajax({
		method: "post",
		url: "cart/ajax_cart_add_item.asp",
		data: {qty: qty, DetailID: DetailID, preorders: comments, ID: wishlist_id, ProductID: ProductID, customorder: customorder}
		})
		.done(function(msg) {
			$('.cart-count').html(parseInt(cart_count) + (parseInt(qty)));
			$('.msg-add-cart-' + wishlist_id).html('<div class="alert alert-success p-1 my-2">Item(s) added to cart</div>');
			setTimeout(function(){
				$('.msg-add-cart-' + wishlist_id).hide();
			}, 2700); // timer wait
			
			setTimeout(function(){
				$('.add-cart-' + wishlist_id).fadeIn('fast');
			}, 3000); // timer wait
		
		})
		.fail(function(msg) {
			console.log("failed");
		})
	}); // end add to cart

	// Open modal to update item
	$('.btn-open-update-item').click(function(e) {
		var wishlist_id = $(this).attr('data-id');
		$('#update-id').val(wishlist_id);
		$('#message-update-item').hide();
		$('#loader-update-item').load("/wishlist/inc-load-wishlist-item-single.asp", {wishlist_id: wishlist_id});	
	});

	// Delete item
	$('#delete-wishlist-item').click(function () {
		var wishlist_id = $('#update-id').val();
		$('#spinner-update-item').show();
		
		$.ajax({
		method: "post",
		url: "wishlist/ajax-wishlist-delete-item.asp",
		data: {wishlist_id: wishlist_id}
		})
		.done(function(msg) {
			$('#spinner-update-item').hide();
			$('#message-update-item').html('<div class="alert alert-success p-2 my-2">Wishlist item has been deleted</span>').fadeIn('slow');
			$('#block-' + wishlist_id).fadeOut('slow');
			$('#updateWishItem').modal('hide');
		})
		.fail(function(msg) {
			$('#spinner-update-item').hide();
			$('#message-update-item').html('<div class="alert alert-danger p-2 my-2">Error deleting item</span>').fadeIn('slow');
		})
	}); // end delete item

	// Update item
	$('#update-wishlist-item').click(function (e) {
		var wishlist_id = $('#update-id').val();
		$('#spinner-update-item').show();
		
		$.ajax({
		method: "post",
		url: "wishlist/ajax-wishlist-update-item.asp",
		data: $('#frm-update-item').serialize() + "&wishlist_id=" + wishlist_id
		})
		.done(function(msg) {
			$('#spinner-update-item').hide();
			$('#message-update-item').html('<div class="alert alert-success p-2 my-2">Wishlist item has been updated.</span>').fadeIn('slow');
			
			$('#priority-' + wishlist_id).html($('#priority').val());
			$('#category-' + wishlist_id).html('Refresh page to see update');
			$('#comments-' + wishlist_id).html($('#comments').val());
			$('#desired-' + wishlist_id).html($('#desired').val());

		})
		.fail(function(msg) {
			$('#spinner-update-item').hide();
			$('#message-update-item').html('<div class="alert alert-danger p-2 my-2">Error updating item</span>').fadeIn('slow');
		})
	}); // end update item
	
	// Add to waiting list
	$('.add-waiting').click(function () {
		var detail_id = $(this).attr('data-detailid');
		var wishlist_id = $(this).attr('data-wishlistid');
		
		$(this).hide();
		$('.waiting-spinner-' + detail_id).show();
		
		$.ajax({
		method: "post",
		url: "products/ajax-add-waiting-list.asp",
		data: {detail_id: detail_id, wishlist_id: wishlist_id}
		})
		.done(function(msg) {
			$('.waiting-spinner-' + detail_id).hide();
			$('.waiting-message-' + detail_id).addClass("alert alert-success").html("Email added to list").fadeIn('slow');
		})
		.fail(function(msg) {
			$(this).show();
			$('.waiting-spinner-' + detail_id).hide();
			$('.waiting-message-' + detail_id).addClass("alert alert-danger").html("Error adding to list").show();
		})
	}); // end add to waiting list
	
	var isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent) ? true : false;
	
	// Load up lists after opening modal
	$('#btn-manage-lists').click(function () {
		$('.category-spinner, .category-message').hide();
		$('.manage-categories').load("wishlist/ajax-wishlist-manage-categories.asp", {retrieve: "retrieve"});		
	}); // end loading up lists
	
	// Add a new category
	$('.btn-add-category').click(function() {
		var category = $("#add-category").val();
		$('.category-spinner').show();

		$.ajax({
		method: "post",
		dataType: "json",
		url: "wishlist/ajax-wishlist-manage-categories.asp",
		data: {category:category, add: "add"}
		})
		.done(function(json, msg) {
			$('#add-category').val('');
			$('.category-message').html('<div class="alert alert-success p-2 my-2">New List Created</div>').fadeIn('slow', function() { $(this).delay(2500).fadeOut('slow'); });
			$('.category-spinner').hide();
			$('.category-message').delay(2500).fadeOut('slow');
			$('.manage-categories').load("wishlist/ajax-wishlist-manage-categories.asp", {retrieve: "retrieve"});
		})
		.fail(function(msg) {
			$('.category-spinner').hide();
			$('.category-message').html('<div class="alert alert-danger p-2 my-2">Error adding category</div>').show();
		})
	}); // end add a new category
	
	// Update category names
	$(document).on('click', '.btn-update-category', function() {
		$('.category-spinner').show();
		var category_id = $(this).attr('data-id');
		var new_name = $('#update-' + category_id).val();
		// console.log("Category ID: " + category_id + "New name: " + new_name);
		$.ajax({
		method: "post",
		url: "wishlist/ajax-wishlist-manage-categories.asp",
		data: {category_id: category_id, new_name: new_name, update: "yes"}
		})
		.done(function(msg) {
			$('.category-spinner').hide();
			$('.category-message').html('<div class="alert alert-success p-2 my-2">Updates have been made! You will need to <a class="alert-link" href="wishlist.asp">refresh this page</a> to have the new names appear in drop-downs.</span>').fadeIn('slow');
		})
		.fail(function(msg) {		
			$('.category-spinner').hide();
			$('.category-message').html('<div class="alert alert-danger p-2 my-2">Error! Only use letters and numbers for wishlist names.</span>').show();
		})
			
	}); // end update category names
	
	
	// Delete a new category
	$(document).on('click', '.delete-category', function() {
		var category_id = $(this).attr('data-id');

		$.ajax({
		method: "post",
		url: "wishlist/ajax-wishlist-manage-categories.asp",
		data: {category_id: category_id, delete: "yes"}
		})
		.done(function(msg) {
			$('#category-' + category_id).fadeOut('slow');
		})
		.fail(function(msg) {
			console.log("failed");
		})
	}); // end delete a new category
	
	
	// Toggle filters for mobile
	$('.link-expand-filters').click(function () {
		$('#expand-filters, .wishlist-filter-down, .wishlist-filter-up').toggle();
	});	
	
	// Share icon show share box
	$('.icon-wishlist-share').click(function () {
		$('.share-box').toggle();
	});	

	new Clipboard('.copy-link'); // Clipboard
	
	// Copy link
	$('.copy-link').click(function() {
		$('.link-copied').show();
		$('.copy-link').hide();
	});	

	// Copy to clipboard does not work on Apple devices or Safair so hide the copy link
	var isApple = /iPad|iPhone/i.test(navigator.userAgent) ? true : false;
	if(isApple) {
		$('.copy-link').hide();
	}	
	// Check for Safari (chrome & safari have both in user agent string)
	if (navigator.userAgent.indexOf('Safari') != -1 && navigator.userAgent.indexOf('Chrome') == -1)
		{
		$('.copy-link').hide();		
		}

// Find broken images and correct image path
$('img').on('error', function (e) {
	img_name = $(this).attr('data-img-name');
	$(this).attr('src', 'http://bodyartforms-products.bodyartforms.com/' + img_name);
  });
