function ResetItemField() {
    $("#scan-item").val('');
    $('#scan-item').prop('readonly', true);
    $("#scan-item").focus();
    $('#scan-item').prop('readonly', false);
};

ResetItemField();



// Load page that shows all items to be pulled
$(document).on("keypress", "#scan-item", function(event){
    if(event.which === 13){ // detect enter keypress

        // Scan barcode and split at . to get purchase order ID and then detail ID
        scanned_item_array = $('#scan-item').val().split('.');
        po_id = scanned_item_array[0];
        scanned_item = scanned_item_array[1];
        $('#load-message').html('');

        $('#load-body').load("restocks/tag-item-info.asp", {item:scanned_item, po_id:po_id});
            ResetItemField();
    }
});

