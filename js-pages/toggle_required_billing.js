 function toggleRequiredBillingTrue() {
		$('#cardNumber, #security-code, #billing-first, #billing-last, #billing-address, #billing-city, #creditCardMonth, #creditCardYear').prop('required', true);
		
	}

 function toggleRequiredBillingFalse() {
		$('#cardNumber, #security-code, #billing-first, #billing-last, #billing-address, #billing-city, #creditCardMonth, #creditCardYear').prop('required', false);
	}		
