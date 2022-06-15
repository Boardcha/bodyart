
// If user from the USA or Canada, show the address validation field
$( document ).ready(function() {
	if(isUserFromUSAorCANADA()){
		if($("#shipping-address").val() !="" && $("#shipping-city").val() !="" && $("#shipping-zip").val() !="" && $("#shipping-country").val() !="" && getShippingState() != ""){
			createAddressBubble("shipping", 'fromCookies');
			$("#shipping-address-container").hide();
		}else{
			showShippingAddressValidation();
			clearShippingAddressInputs();
			$('#selected-shipping-address').hide();	
			$("#shipping-address-container").hide();
		}
	}else{
		hideShippingAddressValidation();
		$("#shipping-address-container").show();
	}
	$('#chk-shipping-manual-address-input').prop('checked', false);
});		
	
function isUserFromUSAorCANADA(){
	var isUserLocal = false;
	$.ajax({
	  dataType: "json",
	  url: "https://pro.ip-api.com/json/?key=1FbdfMEofSriXRs&fields=countryCode",
	  async: false, 
	  success: function(data) {
		if(data.countryCode == 'US' || data.countryCode == 'CA'){
			isUserLocal = true;
		}
	  }
	});	
	return isUserLocal;	
}		

function showShippingAddressValidation(){
	$("#shipping-address-container").hide();	
	$("#shipping-address-autocomplete").show();
	$("#chk-shipping-manual-address-input-container").show();
}

function showBillingAddressValidation(){
	$('#billing-address-autocomplete').show();
	$("#billing-address-container").hide();	
	$("#chk-billing-manual-address-input-container").show();	
}

function hideShippingAddressValidation(){
	$("#shipping-address-autocomplete").hide();
	$("#chk-shipping-manual-address-input-container").hide();
}

function showShippingAddressInputs(){
	$("#shipping-address-container").show();		
	$('#selected-shipping-address').hide();	
	$("#shipping-address-autocomplete").hide();	
	$("#chk-shipping-manual-address-input-container").hide();	
}

$("input[name='chk-shipping-manual-address-input']").change(function(){
	if (this.checked) {
		showShippingAddressInputs();
	}
});	

function closeAddressBubble(section){
	if(section =='billing')
		closeBillingBubble();
	if(section =='shipping')
		closeShippingBubble();			
};

function closeShippingBubble(){
  showShippingAddressValidation();
  clearShippingAddressInputs();
  $('#chk-shipping-manual-address-input').prop('checked', false);
  $('#shipping-full-address').val('');
  $('#shipping-full-address').focus();
};		

function getShippingState(){
	var state = "";
	if($("select[name='shipping-country']").val() == "USA")
		state = $("#shipping-state").val();
	else if($("select[name='shipping-country']").val() == "Canada")
		state = $("select[name='shipping-province-canada']").val();
	else
		state = $("input[name='shipping-province']").val();	
	return state;	
}
	

function createAddressBubble(section, source) {
	var state = "";
	if (section == 'shipping'){
		state = getShippingState();
		hideShippingAddressValidation();
	}	
	if (section == 'billing'){
		state = getBillingState();
		hideBillingAddressValidation();
	}	
	
	var address = '';	
	if(source == 'fromCookies'){
		address = ($('#' + section + '-address').val() + ' ' + $('#' + section + '-address2').val() + '<br/>') +
		($('#' + section + '-city').val() + '<br/>') +
		($('#' + section + '-country').val() + '<br/>') +
		(state + '<br/>') +
		($('#' + section + '-zip').val() + '');
	}else if(source =='fromShipping'){
		address = $("#shipping-address").val() + ' ' + $("#shipping-address2").val() + '<br />' + 
		$("#shipping-city").val() + '<br />' + 
		state + '<br />' + 
		$("#shipping-zip").val() + '<br />' + 
		$("select[name='shipping-country']").val();
	}	
			
	var content = '<div class="alert alert-secondary alert-dismissible fade show" role="alert">' + 
	'  <div id="selected-' + section + '-address-content" class="m-2"><div class="mb-2 font-weight-bold">' + ((section == 'shipping') ? '<i class="fa fa-shipping-fast fa-lg mr-2"></i> ':'') + '<span style="text-transform: uppercase;">' + section + ' ADDRESS</span></div><div style="line-height:22px;">' + address + '</div></div>' +
	'  <button type="button"  class="close" id="btn-edit-' + section + '-address" style="right:20px;padding: 7px 11px 7px 11px;margin-right:16px">' + 
	'	<img src="/images/edit.svg" style="height:14px;width:14px;vertical-align:initial;" />'  +
	'  </button>' +	
	'  <button id="' + section + '-bubble-close" type="button" class="close" data-dismiss="alert" aria-label="Close" style="padding: 7px 11px 7px 11px" onClick="closeAddressBubble(\'' + section + '\')">' + 
	'	<span aria-hidden="true">&times;</span>' +
	'  </button>' +
	'</div>';
	
    $('#selected-' + section + '-address').html(content);
	$('#selected-' + section + '-address').show();	
	$('#' + section + '-country').change();	
	$("input[name='shipping-address2']").val('');
	
}
	

function clearShippingAddressInputs(){
	$("#shipping-full-address").val('');
	$("input[name='shipping-address']").val('');
	$("input[name='shipping-address2']").val('');
	$("input[name='shipping-city']").val('');
	$("input[name='shipping-country']").removeAttr("selected");
	$("input[name='shipping-state']").removeAttr("selected");
	$("input[name='shipping-province-canada']").removeAttr("selected");
	$("input[name='shipping-province']").val('');
	$("input[name='shipping-zip']").val('');
	$('#selected-shipping-address').hide();		
}	

$(document).on('click', '#btn-edit-shipping-address', function() {
	showShippingAddressInputs();	
});
