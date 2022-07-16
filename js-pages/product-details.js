function refreshMenu() {
		var productid = $('#productid').val();
		var gauge = $('#filter-gauge').val();
		var deferred = $.Deferred();

		$('#select-addtocart').load("products/ajax-details-dropdown-addtocart.asp", {productid: productid, gauge: gauge}, function() {
			$('#select-addtocart').show();			
			$('#loading-addtocart').hide();
			deferred.resolve();
		});
		return deferred.promise();
	} 	// end refreshMenu() function
	
	function refreshOptionalColorMenu() {
		var productid = $('#productid').val();
		var deferred = $.Deferred();

		$('#select-anodization').load("products/ajax-anodize-dropdown-addtocart.asp", {productid: productid}, function() {
			$('#select-addtocart').show();			
			$('#loading-addtocart').hide();
			deferred.resolve();
		});
		return deferred.promise();
	} 
	
	function updateSalePrice() {
		var qty = $('input[name="qty"]').val();
		var discount_amount = $('input[name="discount_amount"]').val();
		
		var actual_price = $('.add-cart:checked').attr('data-actual-price');
		var retail_price = $('.add-cart:checked').attr('data-retail-price');
		var sale_price = $('.add-cart:checked').attr('data-sale-price');
		var currency_symbol = $('.add-cart:checked').attr('data-symbol');
		var savings = parseFloat((retail_price - sale_price) * qty).toFixed(2);
		var retail = parseFloat(retail_price * qty).toFixed(2);
		
		if (sale_price != 0 && sale_price != undefined) {
			$('.sale-info').html('<span class="mr-3 ">Savings: ' + currency_symbol + savings + '</span><span class="mr-3">Retail <s>' + currency_symbol + retail + '</s></span>');
		} else {
			$('.sale-info').html('');
		}	
	}

/*
	// Call functions on page load
	refreshMenu().done(function(){
		updateSalePrice();
	}); */
	
	// Filter add to cart menu by gauge
	$('#filter-gauge').change(function(){
		$('#select-addtocart').hide();
		$('#loading-addtocart').show();
		refreshMenu().done(function(){
			updateSalePrice();
		});
		refreshOptionalColorMenu().done(function(){
			updateSalePrice();
		});
	});

	// Never allow qty to go to 0
	$('#add-qty').change(function(){
		var_qty = $("#add-qty").val();
		if (var_qty <= 0) {
			$("#add-qty").val('1');
		}
	});

	

	// Tab switcher -- from bootstrap docs	
	$('#navmenu a').on('click', function (e) {
		e.preventDefault()
		$(this).tab('show')
	  })

	
	// If scripts enabled, switch from submit button to a regular button
	$(".checkout_button").prop("type", "button");

	function resetbutton() {     
	
		$('.add_to_cart').hide();
		$(".add-cart-message").show().html("<div class='alert alert-success my-2 p-1'>Item has been added to your cart</div>");
		$(".add-cart-message").delay(3000).fadeOut("slow");

		$('.add_to_cart').prop("disabled",false);
		$('.add_to_cart').html('<i class="fa fa-shopping-cart fa-lg mr-3"></i>Add to cart');
		$(".add_to_cart").delay(3500).fadeIn("slow", function() {
			$(".add_to_cart").addClass("checkout_button");
			$(".add_to_cart").removeClass("button_loading");
		});	
	}
	
	function showcart() {
		// Display mini shopping cart in corner
		var page_name =  location.pathname.substring(location.pathname.lastIndexOf("/") + 1);
		$(".cart_show").fadeIn('fast');
		$('.cart_show_frame').load("cart/inc_cart_display_mini.asp", {page_name: page_name});
		$(".cart_show").delay(8000).fadeOut('fast');		
	}
	
	function run_cart_mobile() {
		// Do not show mini cart for mobile users, however the inc_cart_main still needs to run to set cart cookies and cart counts	
		var page_name =  location.pathname.substring(location.pathname.lastIndexOf("/") + 1);
		$('.cart_show_frame').load("cart/inc_cart_mobile_nodisplay.asp", {page_name: page_name});
	}

	function changebutton() {   
		// Change styling on button and text
		$('.add_to_cart').prop("disabled",true);
		$('.add_to_cart').html("Adding item to cart <span class='dot-one'>.</span><span class='dot-two'>.</span><span class='dot-three'>.</span>");
		$(".add_to_cart").fadeIn('slow', function() {
			$('.add_to_cart').addClass("button_loading");
			$('.add_to_cart').removeClass("checkout_button");
		});		
	}
	
	var isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent) ? true : false;

	// After button is clicked to add item to cart
	$('.add_to_cart').click(function(e){
        // Fetch form to apply custom Bootstrap validation
				var form = $("#frm-add-cart")
				

        if (form[0].checkValidity() === false) {
            e.preventDefault()
            e.stopPropagation()
						console.log("invalid form elements");
						$(".add-cart-message").show().html("<div class='alert alert-danger my-2 p-1 font-weight-bold'>Quantity &amp; item selection required</div>");
						$(".add-cart-message").delay(3000).fadeOut("slow");
						form[0].classList.add('was-validated');
        } else {
			
					form[0].classList.remove('was-validated');	
				var_preorder6 = ''
				var_preorder7 = ''
				if ($("[name='preorder_field6']").val() != '') {
					var_preorder6 = $.grep([$("[name='preorder_field6_label']").val() + $("[name='preorder_field6']").val()], Boolean).join(", ");
				}
				if ($("[name='preorder_field7']").val() != '') {
					var_preorder7 = $.grep([$("[name='preorder_field7_label']").val() + $("[name='preorder_field7']").val()], Boolean).join(", ");
				}

			changebutton();
			var_detailid = $('.add-cart:checked').val();
			var_anodid = $('.add-anodization:checked').val();
			var_qty = $("[name='qty']").val();
			var_preorders = $.grep([$("textarea#preorders").val(), $("[name='preorder_field1_label']").val() + $("[name='preorder_field1']").val(), $("[name='preorder_field2_label']").val() + $("[name='preorder_field2']").val(), $("[name='preorder_field3_label']").val() + $("[name='preorder_field3']").val(), $("[name='preorder_field4_label']").val() + $("[name='preorder_field4']").val(), $("[name='preorder_field5_label']").val() + $("[name='preorder_field5']").val(), var_preorder6, var_preorder7], Boolean).join(", ");
			if (var_anodid > 0) 
				var_preorders = $('.add-anodization:checked').attr("data-title");
				console.log(var_preorders);
				if (getCookie("cartCount") > 0) {
				var cart_count = $('.cart-count').html();
			} else {
				var cart_count = 0;
			}
			
			var customorder = $("#customorder").val();

				
			if(parseInt($('.add-cart:checked').attr('data-qty')) >= parseInt(var_qty)){
				$.ajax({
				method: "POST",
				url: "cart/ajax_cart_add_item.asp",
				data: {DetailID: var_detailid, qty: var_qty, preorders: var_preorders, customorder: customorder, anodID:var_anodid}
				})
				.done(function( msg ) {
						$(".cart-count").html(parseInt(cart_count) + parseInt(var_qty));
						// Update button to confirm addition to cart
						if(!isMobile) {
							showcart();
						} else {
							run_cart_mobile();
						}
						resetbutton();
						Cookies.set('cartCount', parseInt(cart_count) + parseInt(var_qty), { expires: 365});	
				});
			}else{
				resetbutton();
				$(".add-cart-message").show().html("<div class='alert alert-danger my-2 p-1 font-weight-bold'>We only have " + $('.add-cart:checked').attr('data-qty') + " in stock. Please adjust your quantity.</div>");
			}		
			// Get IP Geo country and region and save to cookie
			$.getJSON("https://pro.ip-api.com/json/?key=1FbdfMEofSriXRs&fields=countryCode,region,city", function() {})
				.done(function( data ) {
					Cookies.set('ip-country', data.countryCode, { expires: 365});
					Cookies.set('ip-region', data.region, { expires: 365});
					Cookies.set('ip-city', data.city, { expires: 365});
				});

			}

			e.preventDefault();
	
		}); // end add to cart click function
		
			
	// Image switcher
	
	// Change cart drop down selector when a color thumbnail is clicked
	// THis code does work, but it's probably not the best usability to change out the drop down so decided to remove it
	
	$(document).on("click", ".img-thumb", function(){
		var img_id = $(this).attr("data-id");
		$('#add-cart-menu label').show();
		$('#msg-filtered-dropdown').html('');

		// only run if there is an actual id to select, and if there are enough items to where the filter gauge drop down is visible
		if(img_id != undefined && img_id != 0) { // && $('#filter-gauge').length > 0 -- to only hide items with long menus
			$('#add-cart-menu label:not(".option_img_' + img_id + '")').hide();
			$('#msg-filtered-dropdown').html('<span class="btn btn-sm btn-info ml-2" id="remove-filter"><i class="fa fa-undo-alt fa-flip-horizontal mr-2"></i>Reload all colors &amp; styles</span>');
		}
	});

	$(document).on("click", "#remove-filter", function(e){
		$('#add-cart-menu label').show();
		$('#msg-filtered-dropdown').html('');
		e.stopPropagation();
	});
	
	
	$(document).on("change", ".add-cart", function(event)
	{
		var imgid = $(this).attr("data-img_id");
		var img_name = $('#img_thumb_' + imgid).attr("data-imgname");
		var data_index = $("#img_id_" + imgid).attr("data-slick-index");
		var selected_text = $(this).attr("dropdown-title");
		//console.log('imgid:' + imgid + ' img_name:' +  img_name + ' data_index:' + data_index);
		// Change out selected drop down text
		$("#selected-item").html(selected_text);
		
		var afterpay_total = parseFloat($('.add-cart:checked').attr('data-actual-price'));
		if(parseFloat($('.add-anodization:checked').attr('data-base-price')) > 0){
			afterpay_total = afterpay_total + parseFloat($('.add-anodization:checked').attr('data-base-price'));
		}
		if(afterpay_total > 0){
			AfterpayWidget(afterpay_total.toString(), minPrice = 0);
		}
		
		if (imgid != 0) {
			$(".slider-main-image").slick('slickGoTo', data_index, true);
		}
	});

	$(document).on("change", ".add-anodization", function(event)
	{
		var selected_text = $(this).attr("dropdown-title");
		$("#selected-anodization").html(selected_text);
		
		var afterpay_total = parseFloat($('.add-cart:checked').attr('data-actual-price'));
		console.log(afterpay_total);
		if(parseFloat($('.add-anodization:checked').attr('data-base-price')) > 0){
			afterpay_total = afterpay_total + parseFloat($('.add-anodization:checked').attr('data-base-price'));
		}
		if(afterpay_total > 0){
			AfterpayWidget(afterpay_total.toString(), minPrice = 0);
		}
		
		console.log(afterpay_total);
	});

	
	$(document).on("click", "#dropdownAddCart", function(){
		$('#add-cart-menu .dropdown-menu').scrollTop(0);
		$('.dropdown-menu').animate({
			scrollTop: 0
	}, 2000);
	});

	$(document).on("click", ".btn-select-menu", function(){
		$("#wishlist-category, #wishlist-priority").prop("selectedIndex", 0);
		$('.wishlist-toggle').hide();
		$('.link-add-wishlist').removeClass('btn-danger');
		$('.link-add-wishlist').addClass('btn-outline-danger');
	});



	// Mobile qty input changer
	$(".qty-add").click(function(){
		var current_value = parseInt($("input[name='qty']").val());
		$("input[name='qty']").val(current_value + 1);	
	});
	
	$(".qty-deduct").click(function(){
		var current_value = parseInt($("input[name='qty']").val());
		var new_value = current_value - 1
		
		if(new_value > 0) {
			$("input[name='qty']").val(current_value - 1);
		}		
	});
	
	
	// Run on page load
		updateSalePrice();

	
	// Change price in add to cart area as select changes
	$(document).on("change", ".add-cart, input[name='qty']", function(){
		updateSalePrice();
	});	

	// Toggle currency menu
	$(".select-currency").on('click', function(event) {
		$('.currency-menu').toggle();
	});		
	
	// Use currency menu to set currency
	$(".currency-menu img").click(function() {
		var selected_currency = $(this).attr("data-currency");
		var selected_symbol = $(this).attr("data-symbol");
		$.ajax({
		type: "post",
		url: "/template/inc-set-currency.asp",
		data: {currency: selected_currency}
		})
		.done(function(msg) {
			$('.ajax-currency').html(selected_symbol + ' ' + selected_currency);
			updateCurrency().done(function(){
				refreshMenu();
				refreshOptionalColorMenu();
			});
		
		});

	});
	
	// Wishlist toggle
	$('.link-add-wishlist').click(function(){
		$('.wishlist-toggle').show();
	});	

	// Threading toggles
	$('.toggle-threading').click(function(){
		$('.div-toggle-threading').fadeToggle();
	});	

	// Add item to wishlist via add to cart form
	$('.link-add-wishlist').on('click', function (event) {
		
		$('.wishlist-message').html('<i class="fa fa-spinner fa-2x fa-spin"></i>').show();
		$("#wishlist-category, #wishlist-priority").prop("selectedIndex", 0);
	
		var_preorders = $.grep([$("textarea#preorders").val(), $("[name='preorder_field1_label']").val() + $("[name='preorder_field1']").val(), $("[name='preorder_field2_label']").val() + $("[name='preorder_field2']").val(), $("[name='preorder_field3_label']").val() + $("[name='preorder_field3']").val(), $("[name='preorder_field4_label']").val() + $("[name='preorder_field4']").val(), $("[name='preorder_field5_label']").val() + $("[name='preorder_field5']").val(), $("[name='preorder_field6_label']").val() + $("[name='preorder_field6']").val(), $("[name='preorder_field7_label']").val() + $("[name='preorder_field7']").val()], Boolean).join(", ");

		$.ajax({
		method: "post",
		url: "/wishlist/ajax-wishlist-add.asp",
		data: $("#frm-add-cart").serialize() + "&preorder_specs=" + encodeURI(var_preorders.replace(/\&/g, ' and '))
		
		})
		.done(function(msg) {
			//console.log($("#frm-add-cart").serialize() + "&preorder_specs=" + var_preorders);
			$('.wishlist-message').html('<div class="alert alert-success p-1 my-2">Item has been added to your wishlist</div>');
			$('.wishlist-message').delay(3000).fadeOut();
			$('.link-add-wishlist').removeClass('btn-outline-danger');
			$('.link-add-wishlist').addClass('btn-danger');
		})
		.fail(function(msg) {
			//console.log($("#frm-add-cart").serialize() + "&preorder_specs=" + var_preorders);
			$('.wishlist-message').html('<div class="alert alert-danger p-1 my-2">Website error. Item did not add to wishlist. Please contact customer service if you continue to have issues.</div>');
			$('.wishlist-message').delay(3000).fadeOut();
		})
	});	// end add item to wishlist

	// Update wishlist item category or priority 
	$('#wishlist-category, #wishlist-priority').on('change', function (event) {
		
		$('.wishlist-message').html('<i class="fa fa-spinner fa-2x fa-spin"></i>').show();
		var category = $('select[name="wishlist-category"]').val();
		var priority = $('select[name="priority"]').val();
	
		$.ajax({
		method: "post",
		url: "/wishlist/ajax-wishlist-update-item.asp",
		data: {category:category, priority:priority, session: "yes"}	
		})
		.done(function(msg) {
			$('.wishlist-message').html('<div class="alert alert-success p-1 my-2">Item has been updated</div>');
			$('.wishlist-message').delay(3000).fadeOut();
		})
		.fail(function(msg) {
			$('.wishlist-message').html('<div class="alert alert-danger p-1 my-2">Website error. Update failed. Please contact customer service if you continue to have issues.</div>');
			$('.wishlist-message').delay(3000).fadeOut();
		})
	});	// end add item to wishlist



	// Add item to wishlist via out of stock items link
	$('.link-wishlist').on('click', function (event) {
		var id = $(this).attr("data-detailid");
		var productid = $(this).attr("data-productid");
		$('.message-outs').hide();
		$('.message-outs-' + id).html('<i class="fa fa-spinner fa-2x fa-spin"></i>').show();
		
		$.ajax({
		method: "post",
		url: "/wishlist/ajax-wishlist-add.asp",
		data: {"add-cart": id, productid: productid}
		})
		.done(function(msg) {
			$('.message-outs-' + id).html('<div class="alert alert-success p-1 my-1">Item has been added to your wishlist</div>');
			$('.message-outs-' + id).delay(3000).fadeOut();
			$('.add-wishlist-' + id).removeClass('btn-outline-danger');
			$('.add-wishlist-' + id).addClass('btn-danger');

		})
		.fail(function(msg) {
			$('.message-outs-' + id).html('<div class="alert alert-danger">Website error. Item did not add to wishlist. Please contact customer service if you continue to have issues.</div>');
			$('.message-outs-' + id).delay(3000).fadeOut();
		})
	});	// end add item to wishlist via out of stock items link
	
	// Get on waiting list - display form when linked clicked
	$('.link-waiting').on('click', function (event) {
		var detailid = $(this).attr("data-detailid");
		
		$('.btn-add-waiting').attr('data-detailid', detailid);
		$('.frm-add-waiting').show().appendTo('.add-waiting-' + detailid);

	});	// end get on waiting list - display form when linked clicked
	
	// Get on waiting list -- submit form
	$('.btn-add-waiting').on('click', function (event) {
		var detailid = $(this).attr("data-detailid");
		$('.frm-add-waiting').hide();
		
		$('.message-outs-' + detailid).html('<i class="fa fa-spinner fa-2x fa-spin"></i>').show();
		
		$.ajax({
		method: "post",
		url: "/products/ajax-add-waiting-list.asp",
		data: $("#frm-add-waiting").serialize()+'&'+$.param({detailid: detailid })
		})
		.done(function(msg) {
			$('.message-outs-' + detailid).html('<div class="alert alert-success my-2 p-1">You have been added to the waiting list</div>');
			$('.message-outs-' + detailid).delay(3000).fadeOut();
		})
		.fail(function(msg) {
			$('.message-outs-' + detailid).html('<div class="notice-red">Website error. Item did not add to the waiting list. Please contact customer service if you continue to have issues.</div>');
			$('.message-outs-' + detailid).delay(3000).fadeOut();
		})
		
	});	// end get on waiting list - submit form

	// FancyBox infinite loop
	$.fancybox.defaults.loop = true;

	// Main photo carousel display
	$('.slider-main-image').slick({
		slidesToShow: 1,
		slidesToScroll: 1,
		fade: true,
		prevArrow: false,
		nextArrow: false /*,
		prevArrow: '<div class="slider-arrow-prev" style="height:10%;top:0;left:0"><i class="fa fa-chevron-left fa-2x text-white pointer border border-primary"></i></div>',
		nextArrow: '<div class="slider-arrow-next" style="height:10%;top:0;left:0"><i class="fa fa-chevron-right fa-2x text-white pointer border border-danger"></i></div>'
		*/
		});

	$('#vert-thumb-carousel').slick({
		asNavFor: '.slider-main-image',
		vertical: true,
		verticalSwiping: true,
		focusOnSelect: true,
		prevArrow: '<i class="fa fa-chevron-up fa-lg p-1 btn-block btn btn-dark btn-sm rounded-0"></i>',
		nextArrow: '<i class="fa fa-chevron-down fa-lg p-1 btn-block btn btn-dark btn-sm rounded-0"></i>',
		responsive: [
			{
				breakpoint: 3000,
				settings: {
				  slidesToShow: 7,
				  slidesToScroll: 1
				}
			  },
			{
		  breakpoint: 1920,
		  settings: {
			slidesToShow: 6,
			slidesToScroll: 1
		  }
		},
		{
		  breakpoint: 1600,
		  settings: {
			slidesToShow: 6,
			slidesToScroll: 1
		  }
		},
		{
		  breakpoint: 1024,
		  settings: {
			slidesToShow: 5,
			slidesToScroll: 1
		  }
		},
		{
		  breakpoint: 600,
		  settings: {
			slidesToShow: 4,
			slidesToScroll: 4
		  }
		},
		{
		  breakpoint: 480,
		  settings: {
			slidesToShow: 3,
			slidesToScroll: 3
		  }
		}
	  ]
	});

	$('#recents,#cross-selling').slick({
		slidesToShow: 6,
		slidesToScroll: 6,
		prevArrow: '<div class="slider-arrow-prev"><i class="fa fa-chevron-circle-left fa-2x text-white pointer"></i></div>',
		nextArrow: '<div class="slider-arrow-next"><i class="fa fa-chevron-circle-right fa-2x text-white pointer"></i></div>',
		responsive: [
		{
		  breakpoint: 3000,
		  settings: {
			slidesToShow: 10,
			slidesToScroll: 10
		  }
		},
		{
		  breakpoint: 1024,
		  settings: {
			slidesToShow: 7,
			slidesToScroll: 7
		  }
		},
		{
		  breakpoint: 600,
		  settings: {
			slidesToShow: 5,
			slidesToScroll: 5
		  }
		},
		{
		  breakpoint: 480,
		  settings: {
			slidesToShow: 3,
			slidesToScroll: 3
		  }
		}
	  ]
	});

function SlickCustomerPhotos() {
	$('#customer-photos').slick({
		slidesToShow: 6,
		slidesToScroll: 6,
		variableWidth: true,
//		infinite: false,
		prevArrow: '<div class="slider-arrow-prev"><i class="fa fa-chevron-circle-left fa-2x text-white pointer"></i></div>',
		nextArrow: '<div class="slider-arrow-next"><i class="fa fa-chevron-circle-right fa-2x text-white pointer"></i></div>',
		responsive: [
		{
		  breakpoint: 1920,
		  settings: {
			slidesToShow: 12,
			slidesToScroll: 10
		  }
		},
		{
		  breakpoint: 1600,
		  settings: {
			slidesToShow: 10,
			slidesToScroll: 8
		  }
		},
		{
		  breakpoint: 1024,
		  settings: {
			slidesToShow: 7,
			slidesToScroll: 6
		  }
		},
		{
		  breakpoint: 600,
		  settings: {
			slidesToShow: 5,
			slidesToScroll: 4
		  }
		},
		{
		  breakpoint: 480,
		  settings: {
			slidesToShow: 3,
			slidesToScroll: 3
		  }
		}
	  ]
	});
} // customer photos function	


	// Automatically load up product photos on page load
	var productid = $("#productid_num").attr("data-productid");
	$.ajax({
		method: "POST",
		url: "/gallery/ajax-photo-gallery.asp",
		data: {productid: productid}
		})
		.done(function(html) {
			response	= html;
			$('#customer-photo-loader').html(response);
			SlickCustomerPhotos(); // re-initialize carousel
		})

		// Filter customer photos by gauge
		$(document).on("change", "#filter_photos_gauge, #filter_photos_color", function()
		{
			var productid = $("#productid_num").attr("data-productid");
			var type = $(this).attr("data-type");
			var filter = $(this).attr("data-filter");
			var gauge = $("#filter_photos_gauge").val();
			var color = $("#filter_photos_color").val();
			var replace= $(this).attr("data-replace");
			var keep = $(this).attr("id");
			var this_value = $(this).val();

			$("#customer-photo-loader").load("/gallery/ajax-photo-gallery.asp?productid=" + productid + "&gauge=" + gauge + "&color=" + color, function() {
				SlickCustomerPhotos(); // re-initialize carousel
			});

			console.log("gauge: " + gauge + " ||  color: " + color + " ||  this value: " + this_value + " ||  type: " + type + " ||  filter: " + filter);

			$.ajax({
				method: "POST",
				url: "/products/ajax-change-select-menus-reviews-gallery.asp",
				data: {productid: productid, type:type, filter:filter, this_value:this_value, gauge:decodeURIComponent(gauge), color: decodeURIComponent(color)}
				})
				.done(function(html) {
					response	= html;
					if (this_value != "All") {
						console.log("refresh other menu");
						$('#' + replace).children('option:not(:selected), optgroup').remove();
						$('#' + replace).append(response); // replace select menu options
					} else {
						console.log("refresh current menu");
						$('#' + keep).children('option:not(:selected), optgroup').remove();
						$('#' + keep).append(response); // replace select menu options
					} // if not set to "all
				});
		});	

	// Page through photos
	$(document).on("click", "#customer-photo-loader .page-link", function(event){
		event.preventDefault();
		var photos_url = $(this).attr("href");
		console.log(photos_url);
		$("#customer-photo-loader").load("/gallery/ajax-photo-gallery.asp" + photos_url, function() {
			SlickCustomerPhotos(); // re-initialize carousel
		});
	});	


		// Automatically load up product reviews on page load
		var productid = $("#productid_num").attr("data-productid");
		$("#div_product_reviews").load("products/ajax_product_reviews.asp?productid=" + productid);	
	
		// Page through reviews
		$(document).on("click", "#div_product_reviews a", function(event)
		{
			event.preventDefault();
			var review_url = $(this).attr("href");
			$("#div_product_reviews").load("products/ajax_product_reviews.asp" + review_url);
		});		
		
		// Filter customer reviews by gauge
		$(document).on("change", "#filter_review_gauge, #filter_review_color", function()
		{
			var productid = $("#productid_num").attr("data-productid");
			var gauge = $("#filter_review_gauge").val();
			var color = $("#filter_review_color").val();
			var replace= $(this).attr("data-replace");
			var keep = $(this).attr("id");
			var this_value = $(this).val();
			var type = $(this).attr("data-type");
			var filter = $(this).attr("data-filter");

			$("#div_product_reviews").load("products/ajax_product_reviews.asp?productid=" + productid + "&gauge=" + gauge + "&color=" + color);

			console.log("gauge: " + gauge + " ||  color: " + color + " ||  this value: " + this_value + " ||  type: " + type + " ||  filter: " + filter);

			$.ajax({
				method: "POST",
				url: "/products/ajax-change-select-menus-reviews-gallery.asp",
				data: {productid: productid, type:type, filter:filter, this_value:this_value, gauge:decodeURIComponent(gauge), color: decodeURIComponent(color)}
				})
				.done(function(html) {
					response	= html;
					if (this_value != "All") {
						console.log("refresh other menu");
						$('#' + replace).children('option:not(:selected), optgroup').remove();
						$('#' + replace).append(response); // replace select menu options
					} else {
						console.log("refresh current menu");
						$('#' + keep).children('option:not(:selected), optgroup').remove();
						$('#' + keep).append(response); // replace select menu options
					} // if not set to "all
				});
		});	
	
	
		// Filter reviews by star rating
		$(document).on("click", ".hover-rating", function()
		{
			var rating = $(this).attr("data-rating");
			var productid = $('#productid').val();
			var gauge = $('#filter_review_gauge').val();
			var color = $("#filter_review_color").val();
			
			$("#div_product_reviews").load("products/ajax_product_reviews.asp?productid=" + productid + "&filter_rating=" + rating + "&gauge=" + gauge + "&color=" + color);
		});	

// Set cookie to show mm sizes in add to cart drop down
$(document).on('click', '#show-mm', function(event) {
	Cookies.set('showmm', "yes", { expires: 365});
	refreshMenu();
	$('#show-mm').hide();
	$('#hide-mm').show();
});

// Removie cookie that shows mm sizes in add to cart drop down
$(document).on('click', '#hide-mm', function(event) {
	Cookies.remove("showmm");
	refreshMenu();
	$('#show-mm').show();
	$('#hide-mm').hide();
});

//Report product photo
$(document).ready(function() {
	$('#modal-report-photo').on('show.bs.modal', function (event) {
	  $(this).find("#report-photo-label").html("Report this photo");
	});

	$("#frmReportPhoto").submit(function(e) {
		var form = $("#frmReportPhoto")

		if (form[0].checkValidity() === false) {
			e.preventDefault()
			e.stopPropagation()
			console.log("invalid form elements");
		} else {  
			var imgSrc = $('.fancybox-slide--current .fancybox-image').attr('src');
			var imgCaption = $('.fancybox-caption').text();
			$('#report-photo-message').load("/emails/ajax_report_photo.asp", {
				comments:  $('#report-photo-comments').val(), url: window.location.href, img_src: imgSrc, caption: imgCaption
			}, function() {
				window.setTimeout(function () {
				  $('#report-photo-message').html("");
				}, 5000 );
			});
		}
		form[0].classList.add('was-validated');
		e.preventDefault();
	});	
});
