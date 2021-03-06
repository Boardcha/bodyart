	// Delete item
	$('.delete-item').click(function () {
		var waiting_id = $(this).attr('data-id');
        $(this).html('<i class="fa fa-spinner fa-lg fa-spin"></i>');

		$.ajax({
		method: "post",
		url: "accounts/ajax-delete-waiting-item.asp",
		data: {waiting_id: waiting_id}
		})
		.done(function(msg) {
			$('#block-' + waiting_id).fadeOut('slow');
		})
		.fail(function(msg) {
			$(this).html('Error deleting item');
		})
	}); // end delete item

	// Update quantity
	$('.update-qty').change(function () {
        var waiting_id = $(this).attr('data-id');
        var waiting_qty = $(this).val();
        $('#msg-update-' + waiting_id).html('<i class="fa fa-spinner fa-lg fa-spin"></i>');
		
		$.ajax({
		method: "post",
		url: "accounts/ajax-waiting-item-update-qty.asp",
		data: {waiting_id: waiting_id, waiting_qty: waiting_qty}
		})
		.done(function(msg) {
            $('#msg-update-' + waiting_id).html('<i class="fa fa-check mr-1"></i>Updated');
            
            setTimeout(function(){
                $('#msg-update-' + waiting_id).html('');
                }, 3000);
		})
		.fail(function(msg) {
			$('#msg-update-' + waiting_id).html('Update failed');
		})
	}); // end update quantity
    
    
    
 // Find broken images and correct image path
$('img').on('error', function (e) {
	img_name = $(this).attr('data-img-name');
	$(this).attr('src', 'http://bodyartforms-products.bodyartforms.com/' + img_name);
  });   