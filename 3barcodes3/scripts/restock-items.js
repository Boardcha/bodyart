function ResetItemField() {
    $("#scan-item, #scan-bin").prop('id', 'scan-item');
    $('#scan-item').prop('placeholder', 'Scan ITEM');
    $("#scan-item").val('');
    $('#scan-item').prop('readonly', true);
    $("#scan-item").focus();
    $('#scan-item').prop('readonly', false);

};

ResetItemField();

var scanned_item = null;
var po_id = null;

// Load page that shows all items to be pulled
$(document).on("keypress", "#scan-item", function(event){
    if(event.which === 13){ // detect enter keypress

        // Scan barcode and split at . to get purchase order ID and then detail ID
        scanned_item_array = $('#scan-item').val().split('.');
        po_id = scanned_item_array[0];
        scanned_item = scanned_item_array[1];
        $('#load-message').html('');

        $('#load-body').load("restocks/restock-item-info.asp", {item:scanned_item, po_id:po_id});

            $(this).val('');
            $(this).prop('id', 'scan-bin');    
            $(this).prop('placeholder', 'Scan BIN');    
            $("#scan-bin").focus();
    }
});

// Display scan match or no scan match
$(document).on("keypress", "#scan-bin", function(event){
    if(event.which === 13){ // detect enter keypress
        scanned_bin = $('#scan-bin').val();
        $('#load-message').html('');
        
		$.ajax({
            method: "post",
            dataType: "json",
            url: "restocks/restock-item-message.asp",
            data: {bin:scanned_bin, item:scanned_item, po_id:po_id}
            })
            .done(function(json, msg) {
                if(json.status == "match"){
                    $('#load-message').html('<div class="alert alert-success">SCAN MATCHED</div>');
                    ResetItemField();
                }
                if(json.status == "no-match"){
                    $('#load-message').html('<div class="alert alert-danger">NO MATCH - WRONG LOCATION</div>');
                    $("#scan-bin").val('');
                    $("#scan-bin").focus();
                }
                
            })
            .fail(function(json, msg) {
                $('#load-message').html('<div class="alert alert-danger">CODE ERROR</div>');
                $("#scan-bin").val('');
                $("#scan-bin").focus();
            })    
    }
});

// Clicking reset button
$(document).on("click", "#btn-reset", function(){ 
    ResetItemField();
    $('#load-message, #load-body').html('');
})

