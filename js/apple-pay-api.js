$( document ).ready(function() {
	
	//Show Apple Pay button if device is compatible
	if (window.ApplePaySession) {
		var merchantIdentifier = 'merchant.com.bodyartforms';
		let promise = ApplePaySession.canMakePaymentsWithActiveCard(merchantIdentifier);
		if (ApplePaySession.canMakePayments(merchantIdentifier)){
			$("#btn-applepay").attr('style', 'display: inline-block !important');
		}else{
			$("#btn-applepay").attr('style', 'display: none !important');
		}
	}else{
		$("#btn-applepay").attr('style', 'display: none !important');
	}

	// Apple Button Click
	const appleButton = document.querySelector('apple-pay-button');
	appleButton.addEventListener('click', function () {
		 
		 var runningTotal	= function() { return parseFloat(subTotal + shippingCost + tax - totalDiscount).toFixed(2); }
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
					shippingCost = data[0].amount;
					selectedShippingId = data[0].identifier;
					selectedShippingCompany = getSelectedShippingCompany(data[0].label);
					shippingOption = data;
				},
				error: function(XMLHttpRequest, textStatus, errorThrown) { 
					console.log("Shipping options error: " + errorThrown); 
				}      	
			}); 
		 }

		 
		 function getSelectedShippingCompany(label){
			if (label.indexOf("USPS") > -1) 
				return "USPS";
			else if (label.indexOf("DHL") > -1) 
				return "DHL";
			else if (label.indexOf("UPS") > -1) 
				return "UPS";
			else
				return "Paid on original order";		 
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
				   shippingCost = parseFloat(json.cost);
				   console.log(shippingCost);
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
		   requiredShippingContactFields: ['postalAddress','email','phone'],
		   requiredBillingContactFields: ['postalAddress'],
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

		function calculateTax(shippingInfo) {
		  
		  tax_country = shippingInfo.countryCode;
		  tax_state = shippingInfo.administrativeArea;
		  tax_zip = shippingInfo.postalCode;

		  var address1 = "";
		  // At this stage, apple only reveals the Country, Locality and the PostCode to protect the privacy of customer at this point. 
		  // They say this is enough for us to determine shipping costs, but not the full address of the customer.
		  // Tax jar api returns a value without address line, too. So, that's true but double check this tax value by comparing with google api

		  
		  $.ajax({
			method: "post",
			dataType: "json",
			async: false,
			url: "/cart/ajax-sales-taxjar-rates.asp",
			data: {initiator: "apple-pay", state_taxed: "yes", shipping_cost: shippingCost, taxable_amount: subTotal, tax_country: tax_country, tax_state: tax_state, tax_zip: tax_zip, tax_address: address1},
			success: function( json ) {
			   tax = parseFloat(json.tax);
			},
			error: function(XMLHttpRequest, textStatus, errorThrown) { 
				console.log("Tax calculation error error: " + errorThrown); 
			}      	
			});  
			return parseFloat(tax).toFixed(2);
		}
		
		session.onshippingcontactselected = function(event) {
			console.log('starting session.onshippingcontactselected');
			console.log('NB: At this stage, apple only reveals the Country, Locality and the PostCode to protect the privacy of what is only a *prospective* customer at this point. This is enough for you to determine shipping costs, but not the full address of the customer.');
			console.log(event);
			
			getShippingOptions( event.shippingContact.countryCode, event.shippingContact.postalCode );
			calculateTax(event.shippingContact);		

			
			var status = ApplePaySession.STATUS_SUCCESS;
			var newShippingMethods = shippingOption;
			var newTotal = { type: 'final', label: '', amount: runningTotal() };
			var newLineItems =[
			{type: 'final',label: subTotalDescr, amount: subTotal }, 
			{type: 'final',label: 'DISCOUNT', amount: totalDiscount },
			{type: 'final',label: 'SHIPPING', amount: shippingCost },
			{type: 'final',label: 'TAX', amount: tax }
			];
			
			session.completeShippingContactSelection(status, newShippingMethods, newTotal, newLineItems )		
			
		}
	
		session.onshippingmethodselected = function(event) {
			
			selectedShippingId = event.shippingMethod.identifier;
			selectedShippingCompany = getSelectedShippingCompany(event.shippingMethod.label);
			shippingCost = getShippingCosts(selectedShippingId);
			
			var status = ApplePaySession.STATUS_SUCCESS;
			var newTotal = { type: 'final', label: '', amount: runningTotal() };
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
			
			var newTotal = { type: 'final', label: '', amount: runningTotal() };
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
			// This is the first stage when we get the "full shipping address" of the customer, in the event.payment.shippingContact object
			console.log(event);

			var promise = sendPaymentToken(event.payment.token, event.payment.shippingContact);
			promise.then(function (success) {	
				var status;
				if (success){
					status = ApplePaySession.STATUS_SUCCESS;
					//document.getElementById("btn-applepay").style.display = "none";
					//document.getElementById("success").style.display = "block";
				} else {
					status = ApplePaySession.STATUS_FAILURE;
				}
				
				session.completePayment(status);
				if (success) window.location = "/checkout_final.asp";
			});
		}

		function sendPaymentToken(paymentToken, shippingInfo) {
			return new Promise(function(resolve, reject) {
			  console.log('starting function sendPaymentToken()');
				
			  // This is where you would pass the payment token to your third-party payment provider to use the token to charge the card. 
			  // Only if your our payment provider tells us the payment was successful should you return a resolve(true) here. Otherwise reject.
			  
			  
			  let jsonData = JSON.stringify(paymentToken.paymentData);
			  var encryptedToken = window.btoa(jsonData);
			  //console.log("payment data:");
			  //console.log(jsonData);			  
	
			  firstName = shippingInfo.givenName;
			  lastName = shippingInfo.familyName;
			  full_name = shippingInfo.givenName + ' ' + shippingInfo.familyName;
			  address1 = shippingInfo.addressLines[0];
			  address2 = shippingInfo.addressLines[1];
			  locality = shippingInfo.locality;
			  administrative_area = shippingInfo.administrativeArea;
			  postal_code = shippingInfo.postalCode;
			  country_code = shippingInfo.countryCode;
			  phone_number = shippingInfo.phoneNumber;
			  email = shippingInfo.emailAddress;
			  amount = subTotal;
			  

			  // START send payment data to authorize.net to process
			  $.ajax({
			  method: "post",
			  //dataType: "json",
			  async: false,
			  url: "checkout/ajax_process_payment.asp",
			  data: {applepay: "on", encryptedToken: encryptedToken, full_name: full_name, address1: address1, address2: address2, locality: locality, 
					 administrative_area: administrative_area, postal_code: postal_code, country_code: country_code, amount: runningTotal(), tax: tax, 
					 shipping_amount: shippingCost, shipping_option: selectedShippingId + "," + shippingCost + "," + selectedShippingCompany,
					 phone_number: phone_number, email: email, first_name: firstName, last_name: lastName
					}
					})
					.done(function( data ) {
						json = JSON.parse(data);
						if (json.stock_status === "fail") {
							console.log("stock_status: fail");
							calcAllTotals();
							alert("Unfortunately we do not have enough quantity in stock for some of the item(s) in your cart.");
							reject;
						}else if (json.flagged === "yes") {
							console.log("ORDER or USER is FLAGGED !!!");
							alert("This order can not be processed online. Please contact customer service for assistance.");							
							reject;
						} else { // If items are in stock 
							if (json.cc_approved === "yes") {
								resolve(true);
								window.location = "/checkout_final.asp";
								console.log("Payment successful");
							} else {				
								console.log("Payment declined");
								alert("Payment declined. " + json.cc_reason);
								reject;
								
							}				
						}			
					})
					.fail(function(xmlHttpRequest, textStatus) {
						alert("Payment declined. Please review your information and try again.");
						reject;
					});			
			});
		}
		
		
		session.oncancel = function(event) {
			console.log('starting session.cancel');
			console.log(event);
		}
		
		session.begin();		
		
	});

});