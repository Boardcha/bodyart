function AfterpayWidget(afterpay_grandtotal, minPrice) {	
	$('.afterpay-paragraph').hide();
	afterpay_grandtotal = afterpay_grandtotal.replace(',','');
	afterpay_amount = afterpay_grandtotal.replace('.','');
	
	if (afterpay_grandtotal >= minPrice) {
		$('#btn-afterpay-checkout').show();
		$('#afterpay-displayonly').hide();
		afterpay_div = '.afterpay-widget';
	} else {
		$('#btn-afterpay-checkout').hide();
		$('#afterpay-displayonly').show();
		afterpay_div = '.afterpay-widget-nonactive';
	}

	//If cart contains gift certificate, hide afterpay button
	setTimeout(function(){
		if($(".cart_item:visible:contains('Digital Gift Certificate')").length > 0 || afterpay_grandtotal > 2000){
			$('#btn-afterpay-checkout').hide();
			$('#afterpay-displayonly').hide();		
		}
	}, 30);

	/* Configure the Widget */
	const apConfig = {
		priceSelector: afterpay_div,
		amount: afterpay_amount,
		locale: 'en_US',
		currency: 'USD',
		minMaxThreshold: {
		min: minPrice * 100,
		max: 200000,
	},
	// variable to remove upper limit
	showUpperLimit: false,
	modalContent: "briogeo",

	};
	/* Initialize the Widget */
	new presentAfterpay(apConfig).init();
}

