 function toggleRequiredBillingTrue() {
		$('#cardNumber, #billing-first, #billing-last, #billing-address, #billing-city, #creditCardMonth, #creditCardYear').prop('required', true);
		
	}

 function toggleRequiredBillingFalse() {
		$('#cardNumber, #billing-first, #billing-last, #billing-address, #billing-city, #creditCardMonth, #creditCardYear').prop('required', false);
	}	