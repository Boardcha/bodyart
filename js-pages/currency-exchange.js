function updateCurrency() {
	var_currency_type = Cookies.get("currency");
	var_show_type = ""
	var_pull_json = 0
	var_symbol = ""
	var currency_deferred = $.Deferred();
	var var_img;
	
	$.ajax({
	type: "post",
	dataType: "json",
	url: "/Connections/openexchange.asp"
	})
	//success: function(json){
	.done(function( json, msg ) {
		
		$('.exchange-price').show();
		$('.usa-price').hide();

		if (var_currency_type=== "USD"){
			var_symbol = "$"
			var_pull_json = json.rates.CAD
			var_show_type = ""
			var_img = 'usa.png'
			$('.row_convert_total').hide();
			$('.afterpay_option').show();
		}
	
	if (var_currency_type=== "CAD"){
		var_symbol = "CAD&nbsp;$"
		var_pull_json = json.rates.CAD
		var_show_type = ""
		var_img = 'canada.png'
		$('.row_convert_total').show();
		$('.afterpay_option').hide();
	}
	
	if (var_currency_type=== "GBP"){
		var_symbol = "£"
		var_pull_json = json.rates.GBP
		var_show_type = ""
		var_img = "uk.png"
		$('.row_convert_total').show();
		$('.afterpay_option').hide();
	}
	
	if (var_currency_type=== "EUR"){
		var_symbol = "€&nbsp;"
		var_pull_json = json.rates.EUR
		var_show_type = ""
		var_img = ""
		$('.row_convert_total').show();
		$('.afterpay_option').hide();
	}
	
	if (var_currency_type=== "AUD"){
		var_symbol = "AUD&nbsp;$"
		var_pull_json = json.rates.AUD
		var_show_type = ""
		var_img = "australia.png"
		$('.row_convert_total').show();
		$('.afterpay_option').hide();
	}
	
	if (var_currency_type=== "NZD"){
		var_symbol = "NZD&nbsp;$"
		var_pull_json = json.rates.NZD
		var_show_type = ""
		var_img = "nz.png"
		$('.row_convert_total').show();
		$('.afterpay_option').hide();
	}

	if (var_currency_type=== "DKK"){
		var_symbol = "kr&nbsp;"
		var_pull_json = json.rates.DKK
		var_show_type = "DKK"
		var_img = "denmark.png"
		$('.row_convert_total').show();
		$('.afterpay_option').hide();
	}
	
	if (var_currency_type=== "JPY"){
		var_symbol = "¥&nbsp;"
		var_pull_json = json.rates.JPY
		var_show_type = "JPY"
		var_img = "japan.png"
		$('.row_convert_total').show();
		$('.afterpay_option').hide();
	}
	
	
	// Write currency type to each span
	$('.currency-type').html(var_show_type);
	if(var_img != "undefined")
	$('#currency-icon').attr('src', '/images/icons/' + var_img);
	
	$('.convert-price').each(function(i, obj) {
		var_newprice = ($(this).attr("data-price") * var_pull_json).toFixed(2);
		$(this).html(var_symbol + var_newprice);
	});

	$.ajax({
		type: "post",
		data: {rate: var_pull_json, symbol:var_symbol, currency: var_currency_type},
		url: "/products/ajax-session-exchange-rate.asp"
		})

	currency_deferred.resolve();
});
	return currency_deferred.promise();

}

