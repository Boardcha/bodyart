$(document).ready(function(){

	var isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry/i.test(navigator.userAgent) ? true : false;

	
if(!isMobile) {
	
   	$('.cart_display_hover').mouseenter(function(){
		$('.cart_show_frame').load("cart/inc_cart_display_mini.asp");
		$('.cart_show').fadeIn("fast");
//		$(".cart_show").animate({ height: 'toggle', opacity: 'toggle' }, 'slow');
	});

	$('#Basket').mouseleave(function(){
		$('.cart_show').fadeOut("fast");
	//	$(".cart_show").animate({ height: 'toggle', opacity: 'toggle' }, 'slow');
	});
}


});