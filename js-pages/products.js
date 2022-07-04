	
	$('.filter-delete').click(function(event) {
	
		var filter = $(this).attr("data-filter");
		var var_value = $(this).attr("data-value");
		console.log(filter + ' : ' + var_value);
	
	//	console.log("Filter: " + filter + "  ||  Value: " + var_value)

		if (filter === "all") {
			// If all filters are being removed
			$('#form-filters input[type="checkbox"]').prop("checked", false);
			$('#form-filters input[type="radio"]').prop("checked", false);
			$('#form-filters input[type="text"]').val('');
			$("#filter-keywords").val('');
		} 
		else if (filter === "keywords") {
			$("#filter-keywords").val('');
			// Do nothing but refresh the form and it will deselect new or restock
		}
		else if (filter === "restock") {
			$("#filter_restock").val('');
		}
		else if (filter === "new") {
			$("#filter_new").val('');
		}
		else if (filter === "stock") {
			$("#filter-stock").val('all');
		}
		else if (filter === "stock-all") {
			$("#filter-stock").val('');
		}
		else if (filter === "price") {
			$("#price-range").val('');
		}		
		else {
			// If specific filter is being removed
			// escape double quote in value (for gauge/length)
			$("input[name='" + filter + "'][value='" + var_value.replace(/"/g, '\"') + "']").prop('checked', false);
		}
	
		
		$('#form-filters').submit();
		return false; // Do nothing on link click
	});

	// START - order by drop down select
	$('input[name="sortby"], input[name="resultsperpage"]').change(function() {
		$('#frm-sort').submit();
	});  // End submit order by
	
	// Scroll to top code
	var offset = 220;
	var duration = 500;
	$(window).scroll(function() {
		if ($(this).scrollTop() > offset) {
			$('.products-top').fadeIn(duration);
		} else {
			$('.products-top').fadeOut(duration);
		}
	});

	// Set display type to list or grid (for mobile only)
	$(document).on("click", ".selector-product-display", function() {
		var display = $(this).attr("data-display");
		if(display === 'list'){
			Cookies.set('product-display', display, { expires: 30});
		} else {
			Cookies.remove('product-display', { path: '/' });
		}
		location.reload();
	});

	
	$('.products-top').click(function(event) {
		event.preventDefault();
		$('html, body').animate({scrollTop: 0}, duration);
		return false;
	});
	
	// Save search to customer account
	$(document).on('click', '.btn-save-search', function(event) {
		var url = $('#save-search-string').val();
		
		$('.btn-save-search').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');
		
		$.ajax({
		method: "post",
		url: "accounts/ajax-save-search.asp",
		data: {url:url}
		})
		.done(function(msg) {
			$('.btn-save-search').html('SEARCH SAVED');	
			$('.btn-save-search').delay(5000).fadeOut('slow');
		})
		.fail(function(msg) {
			$('.btn-save-search').html('<span class="alert alert-danger">SAVE FAILED</span>');
		})
	});  // END save search to customer account

	// Scroll to top of filters if refine results is clicked
	$(document).on('click', '#refine-results', function() {	
		document.getElementById('page-top').scrollIntoView(true);
	});

// filters to auto open

if ($(window).width() > 992) {
	$('#filters').addClass('show');
$('#body-column').addClass('col-lg-9 col-xl-10');
}
//$('#toggle-filters-pc, #toggle-filters-mobile').
// add filters_expanded ="true" and remove the false

// TRACK PINTEREST VIEW PRODUCT CATEGORY PAGE
	pintrk('track', 'viewcategory', {
		line_items: [
		{
		product_category: '<%= title_onpage %>'
		}
		]
	});
