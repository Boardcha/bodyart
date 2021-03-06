/**
 * Define the version of the Google Pay API referenced when creating your
 * configuration
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/request-objects#PaymentDataRequest|apiVersion in PaymentDataRequest}
 */

const environment = "PRODUCTION"; //PRODUCTION, or TEST (used for sandbox)

// ONLY, CHANGE ABOVE LINE to SWITCH BETWEEN PRODUCTION AND LIVE
if (environment == "TEST"){
	var merchantInfo = {merchantName: "Bodyartforms", merchantId: '8265006'};  // This is the sandbox merchantInfo
	var authMerchantId = '483223'; // Authorize.net sandbox
}else{
	var merchantInfo = {merchantName: "Bodyartforms", merchantId: "BCR2DN6T7OWLBOBS"}; //  This is the Production merchantInfo 
	var authMerchantId = '663980'; // Authorize.net production
}	

baseRequest = {
  apiVersion: 2,
  apiVersionMinor: 0
};

/**
 * Card networks supported by your site and your gateway
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/request-objects#CardParameters|CardParameters}
 * @todo confirm card networks supported by your site and gateway
 */
const allowedCardNetworks = ["AMEX", "DISCOVER", "INTERAC", "JCB", "MASTERCARD", "VISA"];

/**
 * Card authentication methods supported by your site and your gateway
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/request-objects#CardParameters|CardParameters}
 * @todo confirm your processor supports Android device tokens for your
 * supported card networks
 */
const allowedCardAuthMethods = ["PAN_ONLY", "CRYPTOGRAM_3DS"];

/**
 * Identify your gateway and your site's gateway merchant identifier
 *
 * The Google Pay API response will return an encrypted payment method capable
 * of being charged by a supported gateway after payer authorization
 *
 * @todo check with your gateway on the parameters to pass
 * @see {@link https://developers.google.com/pay/api/web/reference/request-objects#gateway|PaymentMethodTokenizationSpecification}
 */
const tokenizationSpecification = {
  type: 'PAYMENT_GATEWAY',
  parameters: {
    'gateway': 'authorizenet',
    'gatewayMerchantId': authMerchantId
  }
};

/**
 * Describe your site's support for the CARD payment method and its required
 * fields
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/request-objects#CardParameters|CardParameters}
 */
const baseCardPaymentMethod = {
  type: 'CARD',
  parameters: {
    allowedAuthMethods: allowedCardAuthMethods,
    allowedCardNetworks: allowedCardNetworks
  }
};

/**
 * Describe your site's support for the CARD payment method including optional
 * fields
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/request-objects#CardParameters|CardParameters}
 */
const cardPaymentMethod = Object.assign(
  {},
  baseCardPaymentMethod,
  {
    tokenizationSpecification: tokenizationSpecification
  }
);

/**
 * An initialized google.payments.api.PaymentsClient object or null if not yet set
 *
 * @see {@link getGooglePaymentsClient}
 */
let paymentsClient = null;

/**
 * Configure your site's support for payment methods supported by the Google Pay
 * API.
 *
 * Each member of allowedPaymentMethods should contain only the required fields,
 * allowing reuse of this base request when determining a viewer's ability
 * to pay and later requesting a supported payment method
 *
 * @returns {object} Google Pay API version, payment methods supported by the site
 */
function getGoogleIsReadyToPayRequest() {
  return Object.assign(
      {},
      baseRequest,
      {
        allowedPaymentMethods: [baseCardPaymentMethod]
      }
  );
}

/**
 * Configure support for the Google Pay API
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/request-objects#PaymentDataRequest|PaymentDataRequest}
 * @returns {object} PaymentDataRequest fields
 */
function getGooglePaymentDataRequest() {
  const paymentDataRequest = Object.assign({}, baseRequest);
  paymentDataRequest.allowedPaymentMethods = [cardPaymentMethod];
  paymentDataRequest.transactionInfo = getGoogleTransactionInfo();
  paymentDataRequest.merchantInfo = merchantInfo;
  paymentDataRequest.callbackIntents = ["SHIPPING_ADDRESS",  "SHIPPING_OPTION", "PAYMENT_AUTHORIZATION"];
  paymentDataRequest.shippingAddressRequired = true;
  paymentDataRequest.shippingAddressParameters = getGoogleShippingAddressParameters();
  paymentDataRequest.shippingOptionRequired = true;
  paymentDataRequest.emailRequired = true;

  return paymentDataRequest;
}

/**
 * Return an active PaymentsClient or initialize
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/client#PaymentsClient|PaymentsClient constructor}
 * @returns {google.payments.api.PaymentsClient} Google Pay API client
 */
function getGooglePaymentsClient() {
  if ( paymentsClient === null ) {
    paymentsClient = new google.payments.api.PaymentsClient({
      environment: environment,
      merchantInfo: merchantInfo,
      paymentDataCallbacks: {
        onPaymentAuthorized: onPaymentAuthorized,
        onPaymentDataChanged: onPaymentDataChanged
      }
    });
  }
  return paymentsClient;
}


function onPaymentAuthorized(paymentData) {
  return new Promise(function(resolve, reject){
  // handle the response
  processPayment(paymentData)
    .then(function() {
      resolve({transactionState: 'SUCCESS'});
    })
    .catch(function(error) {
		resolve({
        transactionState: 'ERROR',
        error: {
          intent: 'PAYMENT_AUTHORIZATION',
          message: error,
          reason: 'PAYMENT_DATA_INVALID'
        }
      });
    });
  });
}

/**
 * Handles dynamic buy flow shipping address and shipping options callback intents.
 *
 * @param {object} itermediatePaymentData response from Google Pay API a shipping address or shipping option is selected in the payment sheet.
 * @see {@link https://developers.google.com/pay/api/web/reference/response-objects#IntermediatePaymentData|IntermediatePaymentData object reference}
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/response-objects#PaymentDataRequestUpdate|PaymentDataRequestUpdate}
 * @returns Promise<{object}> Promise of PaymentDataRequestUpdate object to update the payment sheet.
 */
function onPaymentDataChanged(intermediatePaymentData) {

  return new Promise(function(resolve, reject) {
    let shippingAddress = intermediatePaymentData.shippingAddress;
	let address1 = intermediatePaymentData.shippingAddress.address1 + intermediatePaymentData.shippingAddress.address2 + intermediatePaymentData.shippingAddress.address3;
    let shippingOptionData = intermediatePaymentData.shippingOptionData;
	let shippingOptions = getShippingOptions(shippingAddress.countryCode, shippingAddress.postalCode, address1, shippingAddress.administrativeArea, shippingAddress.locality);
    let paymentDataRequestUpdate = {};
    if (intermediatePaymentData.callbackTrigger == "INITIALIZE" || intermediatePaymentData.callbackTrigger == "SHIPPING_ADDRESS") {
	  console.log(getShippingOptions(shippingAddress.countryCode, shippingAddress.postalCode));
      paymentDataRequestUpdate.newShippingOptionParameters = shippingOptions;
	  let defaultShippingOptionId = paymentDataRequestUpdate.newShippingOptionParameters.defaultSelectedOptionId;	  
      paymentDataRequestUpdate.newTransactionInfo = calculateNewTransactionInfo(defaultShippingOptionId, intermediatePaymentData);
	  selectedShippingCompany = getSelectedShippingCompany(getObjectByValue(shippingOptions.shippingOptions, "id", defaultShippingOptionId)[0].label);
	  if(defaultShippingOptionId == "-1")
		 paymentDataRequestUpdate.error =  getGoogleUnserviceableAddressError();
    }
    else if (intermediatePaymentData.callbackTrigger == "SHIPPING_OPTION") {
      paymentDataRequestUpdate.newTransactionInfo = calculateNewTransactionInfo(shippingOptionData.id, intermediatePaymentData);
	  selectedShippingCompany = getSelectedShippingCompany(getObjectByValue(shippingOptions.shippingOptions, "id", shippingOptionData.id)[0].label);
	  if(shippingOptionData.id == "-1")
		 paymentDataRequestUpdate.error =  getGoogleUnserviceableAddressError();	  
    }
	console.log(paymentDataRequestUpdate);
    resolve(paymentDataRequestUpdate);
  });
}

/**
 * Provide Google Pay API with a payment data error.
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/response-objects#PaymentDataError|PaymentDataError}
 * @returns {object} payment data error, suitable for use as error property of PaymentDataRequestUpdate
 */
function getGoogleUnserviceableAddressError() {
	return {
		reason: "SHIPPING_ADDRESS_UNSERVICEABLE",
		message: "Cannot ship to the selected address",
		intent: "SHIPPING_ADDRESS"
	};
}

/**
 * Helper function to create a new TransactionInfo object.

 * @param string shippingOptionId respresenting the selected shipping option in the payment sheet.
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/request-objects#TransactionInfo|TransactionInfo}
 * @returns {object} transaction info, suitable for use as transactionInfo property of PaymentDataRequest
 */
function calculateNewTransactionInfo(shippingOptionId, intermediatePaymentData) {
  let newTransactionInfo = getGoogleTransactionInfo();
	  
  if(totalDiscount > 0.0){
	  newTransactionInfo.displayItems.push({
		type: "LINE_ITEM",
		label: "Discounts",
		price: parseFloat(totalDiscount * -1).toFixed(2),
		status: "FINAL"
	  });  
  }

  selectedShippingId = shippingOptionId;
  shippingCost = getShippingCosts(selectedShippingId);
  
  newTransactionInfo.displayItems.push({
    type: "LINE_ITEM",
    label: "Shipping cost",
    price: shippingCost,
    status: "FINAL"
  });

  if(typeof intermediatePaymentData !== "undefined"){
	  newTransactionInfo.displayItems.push({
		type: "TAX",
		label: "Tax",
		price: calculateTax(intermediatePaymentData, shippingCost),
		status: "FINAL"
	  });  
  }
  
  let totalPrice = 0.00;
  newTransactionInfo.displayItems.forEach(displayItem => totalPrice += parseFloat(displayItem.price));
  newTransactionInfo.totalPrice = parseFloat(totalPrice).toFixed(2).toString();
  totalAmount = parseFloat(totalPrice).toFixed(2);
  return newTransactionInfo;
}

/**
 * Initialize Google PaymentsClient after Google-hosted JavaScript has loaded
 *
 * Display a Google Pay payment button after confirmation of the viewer's
 * ability to pay.
 */
function onGooglePayLoaded() {
  const paymentsClient = getGooglePaymentsClient();
  paymentsClient.isReadyToPay(getGoogleIsReadyToPayRequest())
      .then(function(response) {
        if (response.result) {
          addGooglePayButton();
          // @todo prefetch payment data to improve performance after confirming site functionality
          // prefetchGooglePaymentData();
        }
      })
      .catch(function(err) {
        // show error in developer console for debugging
        console.error(err);
      });
}

/**
 * Add a Google Pay purchase button alongside an existing checkout button
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/request-objects#ButtonOptions|Button options}
 * @see {@link https://developers.google.com/pay/api/web/guides/brand-guidelines|Google Pay brand guidelines}
 */
function addGooglePayButton() {
  const paymentsClient = getGooglePaymentsClient();
  const button = paymentsClient.createButton({
	buttonColor: 'black',
	buttonType: 'checkout',
	buttonSizeMode: 'fill',
	onClick: onGooglePaymentButtonClicked,
	allowedPaymentMethods: [baseCardPaymentMethod]
  });
  document.getElementById('btn-googlepay').appendChild(button);
}
/**
 * Provide Google Pay API with a payment amount, currency, and amount status
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/request-objects#TransactionInfo|TransactionInfo}
 * @returns {object} transaction info, suitable for use as transactionInfo property of PaymentDataRequest
 */
function getGoogleTransactionInfo() {
  return {
    displayItems: [
      {
        label: "Subtotal",
        type: "SUBTOTAL",
        price: parseFloat(subTotal).toString(), // comes from cart_update_totals.js
      }	
	  //, 
      /* Tax will be calculated after the user's address selection, in the function "calculateNewTransactionInfo"
	  {
        label: "Tax",
        type: "TAX",
        price: "0.00",
      }*/
    ],
    countryCode: 'US',
    currencyCode: "USD",
    totalPriceStatus: "FINAL",
    totalPrice: "0.00",
    totalPriceLabel: "Total"
  };
}

/**
 * Calculate tax via taxjar API.
 */
function calculateTax(intermediatePaymentData, shippingCost) {
  
  shipping_cost = shippingCost;
  taxable_amount = subTotal; // comes from cart_update_totals.js
  tax_country = intermediatePaymentData.shippingAddress.countryCode;
  tax_state = intermediatePaymentData.shippingAddress.administrativeArea;
  tax_zip = intermediatePaymentData.shippingAddress.postalCode;
  tax_address =intermediatePaymentData.shippingAddress.address1;
   
  $.ajax({
	method: "post",
	dataType: "json",
	async: false,
	url: "/cart/ajax-sales-taxjar-rates.asp",
	data: {initiator: "google-pay", state_taxed: "yes", shipping_cost: shipping_cost, taxable_amount: taxable_amount, tax_country: tax_country, tax_state: tax_state, tax_zip: tax_zip, tax_address: tax_address},
	success: function( json ) {
       tax = json.tax;
    },
    error: function(XMLHttpRequest, textStatus, errorThrown) { 
        console.log("Tax calculation error error: " + errorThrown); 
    }      	
	});  
	return parseFloat(tax).toFixed(2);
}


/**
 * Provide Google Pay API with shipping address parameters when using dynamic buy flow.
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/request-objects#ShippingAddressParameters|ShippingAddressParameters}
 * @returns {object} shipping address details, suitable for use as shippingAddressParameters property of PaymentDataRequest
 */
function getGoogleShippingAddressParameters() {
  return  {
    phoneNumberRequired: true
  };
}


/**
 * Provide Google Pay API with a payment data error.
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/response-objects#PaymentDataError|PaymentDataError}
 * @returns {object} payment data error, suitable for use as error property of PaymentDataRequestUpdate
 */
function getGoogleUnserviceableAddressError() {
  return {
    reason: "SHIPPING_ADDRESS_UNSERVICEABLE",
    message: "Cannot ship to the selected address",
    intent: "SHIPPING_ADDRESS"
  };
}

/**
 * Prefetch payment data to improve performance
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/client#prefetchPaymentData|prefetchPaymentData()}
 */
function prefetchGooglePaymentData() {
  const paymentDataRequest = getGooglePaymentDataRequest();
  // transactionInfo must be set but does not affect cache
  paymentDataRequest.transactionInfo = {
    totalPriceStatus: 'NOT_CURRENTLY_KNOWN',
    currencyCode: 'USD'
  };
  const paymentsClient = getGooglePaymentsClient();
  paymentsClient.prefetchPaymentData(paymentDataRequest);
}

/**
 * Show Google Pay payment sheet when Google Pay payment button is clicked
 */
function onGooglePaymentButtonClicked() {
  if(preOrderItem == "yes")	{
	checkoutMethod = "btn-googlepay";
    $('#custom-order-warning-modal').modal('show');
	$('#custom-order-items').load("/cart/ajax-pre-order-item-display.asp");	
  }else{	
	const paymentDataRequest = getGooglePaymentDataRequest();
	paymentDataRequest.transactionInfo = getGoogleTransactionInfo();
	const paymentsClient = getGooglePaymentsClient();
	paymentsClient.loadPaymentData(paymentDataRequest);
  }
}

/**
 * Get cookie value by name
*/
function getCookie(cName) {
  const name = cName + "=";
  const cDecoded = decodeURIComponent(document.cookie); 
  const cArr = cDecoded.split('; ');
  let res;
  cArr.forEach(val => {
    if (val.indexOf(name) === 0) res = val.substring(name.length);
  })
  return res
}

var getObjectByValue = function (array, key, value) {
    return array.filter(function (object) {
        return object[key] === value;
    });
};


/**
 * Pull shipping options from DB
*/
function getShippingOptions(countryCode, zipCode, address1, administrativeArea, locality){
	let shippingOptions = [];
	$.ajax({
		method: "post",
		//dataType: "json",
		async: false,
		url: "/google-pay/ajax_display_shipping_options.asp",
		data: {country_code: countryCode, zip_code: zipCode, address: address1, state: administrativeArea, city: locality},
		success: function( data ) {
			data = JSON.parse(data);
			selectedShippingId = data[0].id;
			selectedShippingCompany = getSelectedShippingCompany(data[0].label);
			shippingOptions = data;
		},
		error: function(XMLHttpRequest, textStatus, errorThrown) { 
			console.log("Shipping options error: " + errorThrown); 
		}      	
	}); 

	return {
		defaultSelectedOptionId: shippingOptions[0].id, // First one is the cheapest one
		shippingOptions
	};	
}

/**
 * Get short name for shipping company
*/
function getSelectedShippingCompany(label){
	if (label.indexOf("USPS") > -1) 
		return "USPS";
	else if (label.indexOf("DHL") > -1) 
		return "DHL";
	else if (label.indexOf("UPS") > -1) 
		return "UPS";
	else if (label.indexOf("NO SHIPPING REQUIRED") > -1) 
		return "GIFT CERTIFICATE";		
	else if (label.indexOf("Paid on original order") > -1) 
		return "Paid on original order";			
	else
		return "Undefined";		 
}

/**
 * Get shipping cost for selected shipping option on payment sheet
*/ 
function getShippingCosts(selectedShippingId){
	var shippingCost = 0;
	$.ajax({
		method: "post",
		dataType: "json",
		async: false,
		url: "/google-pay/ajax_get_shipping_cost.asp",
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
		 
/**
 * Process payment data returned by the Google Pay API
 *
 * @param {object} paymentData response from Google Pay API after user approves payment
 * @see {@link https://developers.google.com/pay/api/web/reference/response-objects#PaymentData|PaymentData object reference}
 */
function processPayment(paymentData) {
  return new Promise(function(resolve, reject) {
    setTimeout(function() {
      // @todo pass payment token to your gateway to process payment
      paymentToken = paymentData.paymentMethodData.tokenizationData.token;
	  var encryptedToken = window.btoa(paymentToken);

      full_name = paymentData.shippingAddress.name;
      address1 = paymentData.shippingAddress.address1;
	  address2 = paymentData.shippingAddress.address2;
      locality = paymentData.shippingAddress.locality;
      administrative_area = paymentData.shippingAddress.administrativeArea;
      postal_code = paymentData.shippingAddress.postalCode;
      country_code = paymentData.shippingAddress.countryCode;
      phone_number = paymentData.shippingAddress.phoneNumber;
	  email = paymentData.email;
	  amount = subTotal;

	  $('#btn-googlepay').hide();
	  $('#pay-api-processing-message').show();
	  $('#pay-api-processing-message').html('<div class="alert alert-success mt-2"><i class="fa fa-spinner fa-2x fa-spin"></i> Processing payment ... please wait until payment confirmation screen is displayed.</div>');
		
	  // START send payment data to authorize.net to process
	  $.ajax({
	  method: "post",
	  dataType: "json",
	  async: false,
	  url: "checkout/ajax_process_payment.asp",
	  data: {googlepay: "on", encryptedToken: encryptedToken, full_name: full_name, address1: address1, address2: address2, locality: locality, 
	         administrative_area: administrative_area, postal_code: postal_code, country_code: country_code, amount: totalAmount, tax: tax, 
			 shipping_amount: shippingCost, shipping_option: selectedShippingId + "," + shippingCost + "," + selectedShippingCompany,
			 phone_number: phone_number, email: email
			}
			})
			.done(function( json, msg ) {
				if (json.stock_status === "fail") {
					console.log("stock_status: fail");
					$('#btn-googlepay').show();
					reject("Unfortunately we do not have enough quantity in stock for some of the item(s) in your cart.");
					calcAllTotals();
				}else if (json.flagged === "yes") {
					console.log("ORDER or USER is FLAGGED !!!");
					$('#btn-googlepay').show();
					reject("This order can not be processed online. Please contact customer service for assistance.");							
				} else { // If items are in stock 
					if (json.cc_approved === "yes") {
						resolve({});
						console.log("Payment successful");
						window.location = "/checkout_final.asp";
					} else {				
						console.log("Payment declined");
						console.log("json.errorText: " + json.errorText);
						$('#btn-googlepay').show();
						reject("Payment declined. " + json.cc_reason);
					}				
				}			
			})
			.fail(function() {
				$('#btn-googlepay').show();
				reject("Payment declined. Please review your information and try again.");
			});
    }, 1000);
  });
}