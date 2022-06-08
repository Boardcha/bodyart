
// If user from the USA or Canada, show the address validation field
$( document ).ready(function() {
	if(isUserFromUSAorCANADA()){
		if($("#address").val() !="" && $("#city").val() !="" && $("#zip").val() !="" && $("#country").val() !="" && getShippingState() != ""){
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
	if($("select[name='country']").val() == "USA")
		state = $("#state").val();
	else if($("select[name='country']").val() == "Canada")
		state = $("select[name='province-canada']").val();
	else
		state = $("input[name='province']").val();	
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
		address = ($('#address').val() + ' ' + $('#address2').val() + '<br/>') +
		($('#city').val() + '<br/>') +
		($('#country').val() + '<br/>') +
		(state + '<br/>') +
		($('#zip').val() + '');
	}else if(source =='fromShipping'){
		address = $("#address").val() + ' ' + $("#address2").val() + '<br />' + 
		$("#city").val() + '<br />' + 
		state + '<br />' + 
		$("#zip").val() + '<br />' + 
		$("select[name='country']").val();
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
	$('#country').change();	
	$("input[name='address2']").val('');
	
}
	

function clearShippingAddressInputs(){
	$("#shipping-full-address").val('');
	$("input[name='address']").val('');
	$("input[name='address2']").val('');
	$("input[name='city']").val('');
	$("input[name='country']").removeAttr("selected");
	$("input[name='state']").removeAttr("selected");
	$("input[name='province-canada']").removeAttr("selected");
	$("input[name='province']").val('');
	$("input[name='zip']").val('');
	$('#selected-shipping-address').hide();		
}	

$(document).on('click', '#btn-edit-shipping-address', function() {
	showShippingAddressInputs();	
});
