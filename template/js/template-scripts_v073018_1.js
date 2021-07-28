$(document).ready(function() {

	// Show left-col by default after page loads
	$(".left-col").css("visibility", "visible");
	
	var isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Windows Phone|iemobile|Opera Mini/i.test(navigator.userAgent) ? true : false;
	
	// If iPad (or tablet) is detected then also disable the nav links from using href so that the user can click through them multiple times
	var isIpad = /iPad/i.test(navigator.userAgent) ? true : false;
	if(isIpad) {
		$(".disable-link").click(function(event) {
			event.preventDefault();
		});
	}
	
	
		
	// Needed for Safari
	window.onpopstate = function(event) {
		$('#cart-count').html(Cookies.get("cartCount"));
	};
	
	$('#cart-count').html(Cookies.get("cartCount"));
	
	
	// Remove all empty variables in querystring on any product search using filters
	$('#filter-form').submit(function(){ $('select, input').each(function(){with($(this)) if (val()=='') remove();}) });
	

	// START use jquery-aim plugin to switch through mega menu main nav links with tolerance and not to create a hover tunnel	
	var $menu = $(".mega-nav ul");
	$menu.menuAim({
		tolerance: 0,	// 0 means a lot of tolerance, 100 means less tolerant
		activate: activateSubmenu,
		deactivate: deactivateSubmenu
	});

	function activateSubmenu(row) {
		$('.by-gauge, .new, .by-brand, .plugs').hide();
		var $row = $(row),
			submenuId = $row.data("nav");
			$('.mega-nav ul li').removeClass('nav-selected');
			$(row).addClass('nav-selected');
			$("." + submenuId).show();
	}

	function deactivateSubmenu(row) {
		$('.by-gauge, .new, .by-brand, .plugs').hide();
		var $row = $(row),
			submenuId = $row.data("nav");
		//	$(row).removeClass('nav-selected');
			$('.mega-nav ul li').removeClass('nav-selected');
			$("." + submenuId).hide();
	}	// END use jquery-aim plugin to switch through mega menu main nav links with tolerance and not to create a hover tunnel


	/*
	$("#filter-form").submit(function (e) {
		$('.spinner-window-background, #spinner-viewport').show();
	});
	*/


	// Check all sub boxes for a category
	$('.filter-form .cat-select').on('change',function(){
		var var_name = $(this).attr("data-name");
		$('.sub-' + var_name).prop("checked" , this.checked);
	});
	
	// Navigation section links toggle proper div with mega info
	$(".sub1-nav-link").on('click mouseover', function(event) {	
		$('.col3').hide();
		var nav_type = $(this).attr("data-nav");
		var nav_img = $(this).attr("data-img");
		$('.sub-' + nav_type).fadeIn();
	});

	// Hamburger menu toggle mega-menu for mobile
	$(".hamburger-open").click(function(event) {
		event.preventDefault();	
		var menu_type = $(this).attr("data-type");
		var menu_show = $(this).attr("data-show");
		
		$(".hamburger-open, .hamburger-close").toggle();		
		$("#mega-menu").fadeToggle( "fade", "linear" );
		
		$('.mega-col1').hide();
		$('.' + menu_type + '-down').show();
		$('.mega-nav ul li').removeClass('nav-selected');
		$('.' + menu_show).stop(true, true).show();
		$("#mega-menu").stop(true, true).fadeIn("fast");		
	});	
	// mobile close mega
	$(".hamburger-close").click(function(event) {
		$(".hamburger-open, .hamburger-close").toggle();
		$("#mega-menu").fadeOut( "fade", "linear" );		
	});	

	// Toggle mega-menu
	$(".show-mega").on('click, mouseover', function() {
		var menu_type = $(this).attr("data-type");
		var menu_show = $(this).attr("data-show");

		$('.mega-col1').hide();
		$('.' + menu_type + '-down').show();
		$('.mega-nav ul li').removeClass('nav-selected');
		$('.' + menu_show).stop(true, true).show();
		$("#mega-menu").stop(true, true).fadeIn("fast");
	});

	// Close mega menu on mouseleave 
	$(".top-mega-wrapper, .hide-mega").on('mouseleave', function(event) {
			$("#mega-menu").fadeOut( "fade", "linear" );
	});	
	$(".hide-mega").on('mouseover', function(event) {
			$("#mega-menu").fadeOut( "fade", "linear" );
	});	
	
	// General nav popouts
	$(".nav-open,.nav-wrapper").on('mouseover', function(event) {
		var menu_name = $(this).attr("data-name");
		$('.' + menu_name).show();
		$('.cart_show,#mega-menu').hide();
	});	
	$(".nav-open,.nav-wrapper").on('mouseout', function(event) {	
		var menu_name = $(this).attr("data-name");
		$('.' + menu_name).hide();
		$('.cart_show,#mega-menu').hide();
	});	
	
	// Mob user pop out
	$(".mob-user-icon").on('click', function(event) {
		var menu_name = $(this).attr("data-name");
		$('.' + menu_name).toggle();
		$('#mega-menu').hide();
	});	
	
	
	// Automatically make tooltip fade out on page load
	$('.filters-tooltip').delay(4000).fadeOut('slow');
	// Set cookie if they click hide message
	$(".hide-filter-message").click(function(event) {
		Cookies.set('filters-accessed', 'yes', { expires: 360});
		$('.filters-tooltip').hide();
	});		
	
	// Toggle filters
	$("#icon-filter, #filter-close").click(function(event) {
		$("#left-col").toggleClass('left-col toggle-left-col');
		$("#filter-form").toggleClass('filter-form toggle-filter-form');
		$(".expand-filters").fadeToggle( "fast", "linear" );
		$('#icon-filter').toggle();
		Cookies.set('filters-accessed', 'yes', { expires: 360});
		$('.filters-tooltip').hide();
	});			

	
      // Clear keywords out of search field
      $(".slinput input").keyup(function(){
        var val = $(this).val();   
        if(val.length > 0) {
           $(this).parent().find(".keyword-close-icon").css('color','#555');
        } else {
          $(this).parent().find(".keyword-close-icon").css('color','#ccc');
        }
      });

	// Clear keywords out of search field
	$(".slinput .keyword-close-icon").click(function(){
        $(this).parent().find("input").val('');
        $(this).css('color','#ccc');
      });  
	
	// Allow filters on side to be toggled
	$(".filter-header").click(function(event) {
		var header_name = $(this).attr("data-name");
			$('.toggle-' + header_name + ', .header-' + header_name + ' .fa-angle-down, .header-' + header_name + ' .fa-angle-right').toggle();
	});
	
	
	history.navigationMode = 'compatible';


	// Toggle currency menu
	$(".select-currency").on('click', function(event) {
		$('.currency-menu').toggle();
	});		
	
	// Use currency menu to set currency
	$(".currency-menu li").click(function() {
		var selected_currency = $(this).attr("data-currency");
		var selected_symbol = $(this).attr("data-symbol");
		$.ajax({
		type: "post",
		url: "/template/inc-set-currency.asp",
		data: {currency: selected_currency}
		})
		.done(function(msg) {
			$('.ajax-currency').html(selected_symbol + ' ' + selected_currency)
			location.reload();
		});

	});
	
	/*
	// Show labels on field focusing
	$(".preorder-field").focus(function(){
		var name = $(this).attr("name");
		$("." + name).toggle();
		$(this).data('placeholder',$(this).prop('placeholder'));
		$(this).removeAttr('placeholder')
	});		
	// Show placeholder on field leave
	$(".preorder-field").blur(function(){
		var name = $(this).attr("name");
		$("." + name).toggle();
		$(this).prop('placeholder',$(this).data('placeholder'));
	});		
	*/
	
	
}); // end document.ready