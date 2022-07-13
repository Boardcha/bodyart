function AfterpayWidget() {	
	$('.afterpay-paragraph').hide();
	afterpay_grandtotal = $(".cart_grand-total").html();
	afterpay_grandtotal = afterpay_grandtotal.replace(',','')
	if (afterpay_grandtotal != null) {
			afterpay_amount = afterpay_grandtotal.replace('.','');
			if (afterpay_grandtotal >= 35) {
				$('#btn-afterpay-checkout').show();
				$('#afterpay-displayonly').hide();
				afterpay_div = '.afterpay-widget';
			} else {
				$('#btn-afterpay-checkout').hide();
				$('#afterpay-displayonly').show();
				afterpay_div = '.afterpay-widget-nonactive';
			}
			if (afterpay_grandtotal > 1000) {
				$('#btn-afterpay-checkout').hide();
				$('#afterpay-displayonly').hide();	
			}else{
				$('#btn-afterpay-checkout').show();
				$('#afterpay-displayonly').show();			
			}
	} else {
		afterpay_amount = 0;
		afterpay_div = '.afterpay-widget';
	}
	
	/* Configure the Widget */
	const apConfig = {
		priceSelector: afterpay_div,
		amount: afterpay_amount,
		locale: 'en_US',
		currency: 'USD',
		minMaxThreshold: {
		min: 3500,
		max: 100000,
	},
	// variable to remove upper limit
	showUpperLimit: false,
	modalContent: "briogeo",

	};
	/* Initialize the Widget */
	new presentAfterpay(apConfig).init();
}

AfterpayWidget();