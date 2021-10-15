$( document ).ready(function() {
	
	//Show Apple Pay button if device is compatible
	if (window.ApplePaySession) {
		var merchantIdentifier = 'merchant.com.bodyartforms';
		let promise = ApplePaySession.canMakePaymentsWithActiveCard(merchantIdentifier);
		if (ApplePaySession.canMakePayments(merchantIdentifier)){
			$("apple-pay-button").show();
		}
	}

	// Apple Button Click
	const appleButton = document.querySelector('apple-pay-button');
	appleButton.addEventListener('click', function () {
		 shippingCost = 0; //getShippingCosts('0');
		 var runningTotal	= function() { return subTotal + shippingCost + tax - totalDiscount; }
		 var shippingOption = "";
		 
		 var subTotalDescr	= "SUBTOTAL";
		 
		 function getShippingOptions(countryCode, zipCode){

			$.ajax({
				method: "post",
				dataType: "json",
				async: false,
				url: "/apple-pay/ajax_display_shipping_options.asp",
				data: {country_code: countryCode, zip_code: zipCode},
				success: function( data ) {
				   shippingOption = JSON.parse(data);
				   //shippingOption = [{}];
				   console.log(shippingOption);
				},
				error: function(XMLHttpRequest, textStatus, errorThrown) { 
					shippingOption = [{}];
					console.log("Shipping options error: " + errorThrown); 
				}      	
			}); 
		 }
		 
		 function getShippingCosts(selectedShippingId){
			var shippingCost = 0;
			$.ajax({
				method: "post",
				dataType: "json",
				async: false,
				url: "/apple-pay/ajax_get_shipping_cost.asp",
				data: {shipping_id: selectedShippingId},
				success: function( json ) {
				   shippingCost = json.cost;
				},
				error: function(XMLHttpRequest, textStatus, errorThrown) { 
					console.log("error: " + errorThrown); 
				}      	
			});
			
			return shippingCost;
		 
		 }

		 var paymentRequest = {
		   currencyCode: 'USD',
		   countryCode: 'US',
		   requiredShippingContactFields: ['postalAddress','email'],
		   //requiredBillingContactFields: ['postalAddress','email', 'name', 'phone'],
		   lineItems: [
		   {label: subTotalDescr, amount: subTotal }, 
		   {label: 'DISCOUNT', amount: totalDiscount },
		   {label: 'SHIPPING', amount: shippingCost },
		   {label: 'TAX', amount: tax }
		   ],
		   total: {
			  label: 'GRAND TOTAL',
			  amount: runningTotal()
		   },
		   supportedNetworks: ['masterCard', 'visa', 'amex', 'discover'],
		   merchantCapabilities: [ 'supports3DS', 'supportsEMV', 'supportsCredit', 'supportsDebit' ]
		};
		
		var session = new ApplePaySession(1, paymentRequest);
		
		// Merchant Validation
		session.onvalidatemerchant = function (event) {
			console.log(event);
			var promise = performValidation(event.validationURL);
			promise.then(function (merchantSession) {
				session.completeMerchantValidation(merchantSession);
			}); 
		}
		

		function performValidation(valURL) {
			return new Promise(function(resolve, reject) {
				var xhr = new XMLHttpRequest();
				xhr.onload = function() {
					var data = JSON.parse(this.responseText);
					console.log(data);
					resolve(data);
				};
				xhr.onerror = reject;
				xhr.open('GET', '/apple-pay/applepay-merchant-validation.asp?url=' + valURL);
				xhr.send();
			});
		}

		session.onshippingcontactselected = function(event) {
			console.log('starting session.onshippingcontactselected');
			console.log('NB: At this stage, apple only reveals the Country, Locality and 4 characters of the PostCode to protect the privacy of what is only a *prospective* customer at this point. This is enough for you to determine shipping costs, but not the full address of the customer.');
			console.log(event);
			
			getShippingOptions( event.shippingContact.countryCode, event.shippingContact.postalCode );
			
			// UPDATE TAX VARIABLE HERE
			
			var status = ApplePaySession.STATUS_SUCCESS;
			var newShippingMethods = shippingOption;
			var newTotal = { type: 'final', label: 'Bodyartforms', amount: runningTotal() };
			var newLineItems =[
			{type: 'final',label: subTotalDescr, amount: subTotal }, 
			{type: 'final',label: 'DISCOUNT', amount: totalDiscount },
			{type: 'final',label: 'SHIPPING', amount: shippingCost },
			{type: 'final',label: 'TAX', amount: tax }
			];
			
			session.completeShippingContactSelection(status, newShippingMethods, newTotal, newLineItems );
			
			
		}
		
		session.onshippingmethodselected = function(event) {
			console.log('starting session.onshippingmethodselected');
			console.log(event);
			
			getShippingCosts( event.shippingMethod.identifier);
			
			var status = ApplePaySession.STATUS_SUCCESS;
			var newTotal = { type: 'final', label: 'Bodyartforms', amount: runningTotal() };
			var newLineItems =[
			{type: 'final',label: subTotalDescr, amount: subTotal }, 
			{type: 'final',label: 'DISCOUNT', amount: totalDiscount },
			{type: 'final',label: 'SHIPPING', amount: shippingCost },
			{type: 'final',label: 'TAX', amount: tax }
			];
			
			session.completeShippingMethodSelection(status, newTotal, newLineItems );
			
			
		}
		
		session.onpaymentmethodselected = function(event) {
			console.log('starting session.onpaymentmethodselected');
			console.log(event);
			
			var newTotal = { type: 'final', label: 'Bodyartforms', amount: runningTotal() };
			var newLineItems =[
			{type: 'final',label: subTotalDescr, amount: subTotal }, 
			{type: 'final',label: 'DISCOUNT', amount: totalDiscount },
			{type: 'final',label: 'SHIPPING', amount: shippingCost },
			{type: 'final',label: 'TAX', amount: tax }
			];
			
			session.completePaymentMethodSelection( newTotal, newLineItems );
			
			
		}
		
		session.onpaymentauthorized = function (event) {

			console.log('starting session.onpaymentauthorized');
			console.log('NB: This is the first stage when you get the *full shipping address* of the customer, in the event.payment.shippingContact object');
			console.log(event);

			var promise = sendPaymentToken(event.payment.token);
			promise.then(function (success) {	
				var status;
				if (success){
					status = ApplePaySession.STATUS_SUCCESS;
					document.getElementById("applePay").style.display = "none";
					document.getElementById("success").style.display = "block";
				} else {
					status = ApplePaySession.STATUS_FAILURE;
				}
				
				console.log( "result of sendPaymentToken() function =  " + success );
				session.completePayment(status);
			});
		}

		function sendPaymentToken(paymentToken) {
			return new Promise(function(resolve, reject) {
				console.log('starting function sendPaymentToken()');
				console.log(paymentToken);
				
				console.log("this is where you would pass the payment token to your third-party payment provider to use the token to charge the card. Only if your provider tells you the payment was successful should you return a resolve(true) here. Otherwise reject;");
				console.log("defaulting to resolve(true) here, just to show what a successfully completed transaction flow looks like");
				if ( debug == true )
				resolve(true);
				else
				reject;
			});
		}
		
		session.oncancel = function(event) {
			console.log('starting session.cancel');
			console.log(event);
		}
		
		session.begin();		
		
	});

	
});