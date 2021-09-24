/**
 * Define the version of the Google Pay API referenced when creating your
 * configuration
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/request-objects#PaymentDataRequest|apiVersion in PaymentDataRequest}
 */

var tax = 0.0;
var shippingCost = 0.0;
var totalAmount = 0.0;
var selectedShippingId = 0;
const baseRequest = {
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
    'gatewayMerchantId': '483223' // sandbox: 483223 / production: 663980
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
  paymentDataRequest.merchantInfo = {
    // @todo a merchant ID is available for a production environment after approval by Google
    // See {@link https://developers.google.com/pay/api/web/guides/test-and-deploy/integration-checklist|Integration checklist}
    merchantId: '8265006',
    merchantName: 'Bodyartforms'
  };

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
      environment: "TEST",
      merchantInfo: {
        merchantName: "Bodyartforms",
        merchantId: "8265006"
      },
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
    .catch(function() {
        resolve({
        transactionState: 'ERROR',
        error: {
          intent: 'PAYMENT_AUTHORIZATION',
          message: 'Insufficient funds',
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
    let shippingOptionData = intermediatePaymentData.shippingOptionData;
    let paymentDataRequestUpdate = {};

    if (intermediatePaymentData.callbackTrigger == "INITIALIZE" || intermediatePaymentData.callbackTrigger == "SHIPPING_ADDRESS") {
	  //========== ZIP CODES THAT DHL DOES NOT DELIVER TO AND NEED TO BE FORCED TO USPS ========
	  const notAvailableForDHL = ["96799", "96910", "96912", "96913", "96915", "96916", "96917", "96919", "96921", "96923", "96928", "96929", "96931", "96932", "96939", "96940", "96941", "96942", "96943", "96944", "96950", "96951", "96952", "96960", "96970"];
	  if(notAvailableForDHL.includes(shippingAddress.postalCode)){
        paymentDataRequestUpdate.newShippingOptionParameters = getShippingOptions(shippingAddress.countryCode, true); // USPSOnly = true
	  }else {
		paymentDataRequestUpdate.newShippingOptionParameters = getShippingOptions(shippingAddress.countryCode, false); // USPSOnly = false
	  }
      let selectedShippingOptionId = paymentDataRequestUpdate.newShippingOptionParameters.defaultSelectedOptionId;
      paymentDataRequestUpdate.newTransactionInfo = calculateNewTransactionInfo(selectedShippingOptionId, intermediatePaymentData);
    }
    else if (intermediatePaymentData.callbackTrigger == "SHIPPING_OPTION") {
      paymentDataRequestUpdate.newTransactionInfo = calculateNewTransactionInfo(shippingOptionData.id, intermediatePaymentData);
    }
    resolve(paymentDataRequestUpdate);
  });
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

  shippingCost = getShippingCosts()[shippingOptionId];
  selectedShippingId = shippingOptionId;
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
  console.log("grandTotal:" + totalAmount.toString());
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
	buttonColor: 'white',
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
        price: parseFloat(totalWithoutShipping).toString(), // comes from cart_update_totals.js
      }//, 
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
  
  shipping_cost = shippingCost
  taxable_amount = totalWithoutShipping // comes from cart_update_totals.js
  tax_country = intermediatePaymentData.shippingAddress.countryCode
  tax_state = intermediatePaymentData.shippingAddress.administrativeArea
  tax_zip = intermediatePaymentData.shippingAddress.postalCode
  tax_address =intermediatePaymentData.shippingAddress.address1
   
  $.ajax({
	method: "post",
	dataType: "json",
	async: false,
	url: "cart/ajax-sales-taxjar-rates.asp",
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
 * Provide a key value store for shippping options.
 */
function getShippingCosts() {
  return {
	"0": "0.00",   // Do not charge shipping amount on addons
    "10": "13.95", //USPS Priority mail heavy
    "25": "23.95", //USPS Express mail
    "30": "4.95",  //USPS First Class Mail
	"31": "7.95",  //USPS Priority mail
	"3": "4.95",   //DHL Expedited Max
	"7": "0.00",   //Free: DHL Basic mail
	"26": "44.95", //USPS Express mail international
	"11": "2.95",  //DHL GlobalMail Packet Priority
	"28": "4.95",  //DHL GlobalMail Parcel Priority
	"14": "31.95", //USPS Global priority mail
	"27": "65.95", //USPS Express mail international
	"13": "2.95",  //DHL GlobalMail Packet Priority
	"29": "5.95"  //DHL GlobalMail Parcel Priority
  }
}

/**
 * Provide a key value store for shippping options.
 */
function getShippingCompany() {
  return {
	"0": "Paid on original order", // Do not charge shipping amount on addons
    "10": "USPS", //USPS Priority mail heavy
    "25": "USPS", //USPS Express mail
    "30": "USPS",  //USPS First Class Mail
	"31": "USPS",  //USPS Priority mail
	"3": "DHL",   //DHL Expedited Max
	"7": "DHL",   //Free: DHL Basic mail
	"26": "USPS", //USPS Express mail international
	"11": "DHL",  //DHL GlobalMail Packet Priority
	"28": "DHL",  //DHL GlobalMail Parcel Priority
	"14": "USPS", //USPS Global priority mail
	"27": "USPS", //USPS Express mail international
	"13": "DHL",  //DHL GlobalMail Packet Priority
	"29": "DHL"  //DHL GlobalMail Parcel Priority
  }
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
 * Provide Google Pay API with shipping options and a default selected shipping option.
 *
 * @see {@link https://developers.google.com/pay/api/web/reference/request-objects#ShippingOptionParameters|ShippingOptionParameters}
 * @returns {object} shipping option parameters, suitable for use as shippingOptionParameters property of PaymentDataRequest
 */
function getShippingOptions(countryCode, USPSOnly) {

	let shippingOptions = [];
	let addons = getCookie("OrderAddonsActive");
	console.log("cookie addons:" + addons);
	if (addons !== "" && typeof addons !== "undefined" ) 
		shippingOptions.push({
			"id": "0",
			"label": "Free: Paid on original order",
			"description": ""
		});	
		
	if (countryCode =='US' && USPSOnly == false && shippingWeight <= 8) 
		shippingOptions.push({
			"id": "7",
			"label": "Free: DHL Basic mail",
			"description": "Average delivery time 7-14 business days."
		});	

	if (countryCode =='CA' && USPSOnly == false && shippingWeight <= 15) 
		shippingOptions.push({
			"id": "11",
			"label": "$2.95: DHL GlobalMail Packet Priority",
			"description": "TRACKED TO USA BORDER ONLY. Average delivery time is 3-4 weeks."
		  });
		  
	if (countryCode =='US' && USPSOnly == false && shippingWeight <= 32) 
		shippingOptions.push({
			"id": "3",
			"label": "$4.95: DHL Expedited Max",
			"description": "Estimated delivery time: 3-4 business days."
		});
		
	if (countryCode =='US' && shippingWeight <= 3) 
		shippingOptions.push({
			"id": "30",
			"label": "$4.95: USPS First Class Mail",
			"description": "Average delivery time 7-14 business days."
		});

	if (countryCode =='US' && shippingWeight <= 32) 
		shippingOptions.push({
			"id": "31",
			"label": "$7.95: USPS Priority mail",
			"description": "Estimated delivery time: 2-3 business days."
		});
		
	if (countryCode =='US') // HEAVY
		shippingOptions.push({
			"id": "10", 
			"label": "$13.95: USPS Priority mail heavy", 
			"description": "Estimated delivery time: 2-3 business days."
		});
		
	if (countryCode =='US' && shippingWeight <= 32) 
		shippingOptions.push({
			"id": "25",
			"label": "$23.95: USPS Express mail",
			"description": "Estimated delivery time: 1-2 business days."
		});

	if (countryCode =='CA' && USPSOnly == false) // HEAVY 
		shippingOptions.push({
			"id": "28",
			"label": "$4.95: DHL GlobalMail Parcel Priority",
			"description": "FULLY TRACKABLE. Average delivery time is 3-4 weeks."
		});	
		
	if (countryCode =='CA' && shippingWeight <= 16) 
		shippingOptions.push({
			"id": "26",
			"label": "$44.95: USPS Express mail international",
			"description": "Estimated delivery time 3-5 days."
		});			
		
	if (countryCode != 'CA' && countryCode != 'US' && USPSOnly == false && shippingWeight <= 15) // INTERNATIONAL
		shippingOptions.push({
			"id": "13",
			"label": "$2.95: DHL GlobalMail Packet Priority",
			"description": "TRACKED TO USA BORDER ONLY. Average delivery time is 3-4 weeks."
		});	

	if (countryCode != 'CA' && countryCode != 'US' && USPSOnly == false) // INTERNATIONAL / HEAVY
		shippingOptions.push({
			"id": "29",
			"label": "$5.95: DHL GlobalMail Parcel Priority",
			"description": "FULLY TRACKABLE. Average delivery time is 3-4 weeks."
		});	
		
	if (countryCode != 'CA' && countryCode != 'US' && shippingWeight <= 13) // INTERNATIONAL
		shippingOptions.push({
			"id": "14",
			"label": "$31.95: USPS Global priority mail",
			"description": "Average delivery time is 3-4 weeks."
		});		

	if (countryCode != 'CA' && countryCode != 'US' && shippingWeight <= 16) // INTERNATIONAL
		shippingOptions.push({
			"id": "27",
			"label": "$65.95: USPS Express mail international",
			"description": "Estimated delivery time 3-5 business days."
		});	
	
	return {
		defaultSelectedOptionId: shippingOptions[0].id, // First one is the cheapest one
		shippingOptions
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
  const paymentDataRequest = getGooglePaymentDataRequest();
  paymentDataRequest.transactionInfo = getGoogleTransactionInfo();

  const paymentsClient = getGooglePaymentsClient();
  paymentsClient.loadPaymentData(paymentDataRequest);
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

/**
 * Process payment data returned by the Google Pay API
 *
 * @param {object} paymentData response from Google Pay API after user approves payment
 * @see {@link https://developers.google.com/pay/api/web/reference/response-objects#PaymentData|PaymentData object reference}
 */
function processPayment(paymentData) {

  return new Promise(function(resolve, reject) {
    setTimeout(function() {
      // show returned data in developer console for debugging
      console.log(paymentData);
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
	  amount = totalWithoutShipping

	  // START send payment data to authorize.net to process
	  $.ajax({
	  method: "post",
	  dataType: "json",
	  async: false,
	  url: "checkout/ajax_process_payment.asp",
	  data: {googlepay: "on", encryptedToken: encryptedToken, full_name: full_name, address1: address1, address2: address2, locality: locality, 
	         administrative_area: administrative_area, postal_code: postal_code, country_code: country_code, amount: totalAmount, tax: tax, 
			 shipping_amount: shippingCost, shipping_option: selectedShippingId + "," + shippingCost + "," + getShippingCompany()[selectedShippingId],
			 phone_number: phone_number, email: email
			}
			})
			.done(function( json, msg ) {
				if (json.stock_status === "fail") {
					console.log("stock_status: fail");
					reject(new Error('Unfortunately we do not have enough quantity in stock for some of the item(s) in your cart.'));
					calcAllTotals();
				}else if (json.flagged === "yes") {
					console.log("Order or user is flagged.");
					reject(new Error('This order can not be processed online. Please contact customer service for assistance.'));				
				} else { // If items are in stock 
					if (json.cc_approved === "yes") {
						resolve({});
						console.log("Payment successful");
						window.location = "/checkout_final.asp";
					} else {				
						console.log("Payment declined");
						console.log("msg.responseText: " + msg.responseText);
						console.log("json.errorText: " + json.errorText);
						reject(new Error('Payment declined. ' + json.cc_reason));
					}				
				}			
			})
			.fail(function() {
				reject(new Error('Payment declined. Please review your information and try again.'));
			});
    }, 3000);
  });
}