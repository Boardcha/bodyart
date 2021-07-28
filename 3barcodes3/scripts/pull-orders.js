function ResetItemField() {
    $("#scan-invoice, #scan-item").prop('id', 'scan-invoice');
    $('#scan-invoice').prop('placeholder', 'Scan INVOICE');
    $("#scan-invoice").val('');
    $('#scan-invoice').prop('readonly', true);
    $("#scan-invoice").focus();
    $('#scan-invoice').prop('readonly', false);

};

ResetItemField();

// Clicking reset button
$(document).on("click", "#btn-reset", function(){ 
    ResetItemField();
    $('#load-message, #load-body').html('');
})

// Load page that shows all items to be pulled
$(document).on("keypress", "#scan-invoice", function(event){
    if(event.which === 13){ // detect enter keypress
    $('#load-message').html('');
    var scanned_invoicenum = $('#scan-invoice').val();

    $('#load-body').load("pulling/items-to-pull.asp", {invoices:scanned_invoicenum});
    console.log(event.type + ": " +  event.which);
    if(event.which === 13){ // detect enter keypress after page load as well
        console.log("enter key pressed");
        //stuff goes here after page load
        //$('#btn-load-items').hide();
        // Set form field focus
        $(this).val('');
        $(this).prop('id', 'scan-item');    
        $(this).prop('placeholder', 'Scan ITEM');    
        $("#scan-item").focus();
        $('#sort-section').show();
        $('#display-invoice').html(scanned_invoicenum);
    }
    } // detect enter key on main invoice scan 1st
})

$(document).on("keypress", "#scan-item", function(event){
    if(event.which === 13){ // detect enter keypress
    $('#load-message').html('');
        $('#scan-item').prop('readonly', true);
        $("#scan-item").focus();
        $('#scan-item').prop('readonly', false);
    var invoiceid = $('#display-invoice').html();

    // Find first item in list with scanned matching productdetailid 



        if($('#scan-item').val().indexOf(".") != -1){
            // . period found, only use the number after the period
            scanned_item_array = $('#scan-item').val().split('.');
            scanned_number = scanned_item_array[1];
            console.log(". found")
        } else {
            // . period not found, default to regular item scan
            var scanned_number = $('#scan-item').val();
        }

    $( "#scan-item" ).val('');
    var row_found = $("tr[data-location='" + scanned_number + "']:first").attr('data-orderdetailid');
   

    if(row_found != undefined) {
        row_found = $.trim(row_found);
        // Hide previous found rows
        $('.done').hide();

        // Update the amount of times scanned
        $.ajax({
            method: "post",
            url: "pulling/set-times-scanned.asp",
            data: {orderdetailid: row_found}
            })

       // If row is done, then display too many scan error, otherwise continue
        if ($('#' + row_found).hasClass('done')) {
            $('#load-message').html('<div class="alert alert-danger font-weight-bold">TOO MANY SCANS &#9888; You should only have ' + $('#' + row_found).attr('data-qty') + '</div>')
            $('html, body').animate({
                scrollTop: $('#scan-item').offset().top
            }, 0);
        } else {

        //Highlight and move screen to top of row
        $('#' + row_found + ', #' + row_found + '_sub').addClass('table-warning');
        $('html, body').animate({
            scrollTop: $('#' + row_found).offset().top
        }, 0);
        $('#' + row_found + '_sub').show();

        // increment the scans attribute
        var times_scanned = parseInt($('#' + row_found).attr('data-timescanned'));
        var orig_qty = parseInt($('#' + row_found).attr('data-qty'));
        var row_match_qty = $('#' + row_found).attr('data-matchqty');

        if (row_match_qty === 'no') {
            // Most items will need incrementing
            var incrememt_scan = times_scanned + 1;
        } else {
            // Items like orings that we don't want to scan a zillion times
            var incrememt_scan = orig_qty;
        }
        $('#' + row_found).attr('data-timescanned', incrememt_scan);
        //var qty_to_pull = parseInt($('#' + row_found).attr('data-qty'));
        $('#still_need_' + row_found).html(incrememt_scan);
        
        // If the amount scanned and qty are equal set the row to a finished status to be hidden next time a scan is made
        if($('#' + row_found).attr('data-timescanned') === $('#' + row_found).attr('data-qty')) {
            $('#' + row_found + ', #' + row_found + '_sub').addClass('done');
            $('#' + row_found + ', #' + row_found + '_sub').removeClass('table-warning');
            $('#' + row_found + ', #' + row_found + '_sub').addClass('table-secondary');
            $('#check_' + row_found).addClass('complete text-success');
        }
        } // if has class done is not found, display too many scans error
    } 
    // if item not found, then either it's a no scan match OR an overscan
    else { 
        $('#load-message').html('<div class="alert alert-danger font-weight-bold">NO SCAN MATCH</div>')
        $('html, body').animate({
            scrollTop: $('#scan-item').offset().top
        }, 0);
    }

    // Order complete message
    var checkmarks = $(".scan_complete:not(.complete)").length;
    if ( checkmarks > 0) {
        // Items still need to be scanned, order is not complete
         
    } else {
        ResetItemField();
        $('#load-message').html('<div class="alert alert-success font-weight-bold h5">ORDER COMPLETE</div>')
        $('html, body').animate({
            scrollTop: $('#scan-item').offset().top
        }, 0);

        $.ajax({
            method: "post",
            url: "pulling/set-timestamps.asp",
            data: {invoiceid: invoiceid}
            })
    }
    
} // end enter keypress
})

// Expand row when tapping picture
$(document).on("click", ".expand", function(){
    var orderdetailid = $(this).attr('data-orderdetailid');
    $('#' + orderdetailid + '_sub').toggle();

})

// Toggle done checkmark
$(document).on("click", ".toggle-done", function(){
    var orderdetailid = $(this).attr('data-orderdetailid');
    $('#' + orderdetailid + ', #' + orderdetailid + '_sub').toggleClass('done');
    $('#' + orderdetailid + ', #' + orderdetailid + '_sub').toggleClass('table-secondary');
    $('#check_' + orderdetailid).toggleClass('complete');
    $('#check_' + orderdetailid).toggleClass('text-success');

})


// Copy orderdetailid to attribute for backorder button
$(document).on("click", ".bo-button", function(){
    var orderdetailid = $(this).attr('data-orderdetailid');
    $('#btn-submit-bo').attr('data-orderdetailid', orderdetailid);

    $('#new-bo-message').html('');
    $('#btn-submit-bo, #frm-backorder').show();
})


// Backorder item
$(document).on("click", "#btn-submit-bo", function(){
    var notes = $('input[name=entireorder]:checked').val() + ' ' + $('#partial option:selected').val();
    var orderdetailid = $(this).attr('data-orderdetailid');

      $.ajax({
		method: "post",
		url: "pulling/backorder-item.asp",
		data: {notes: notes, orderdetailid: orderdetailid}
		})
		.done(function(msg) {
            $('#new-bo-message').html('<div class="alert alert-success">ITEM SUCCESSFULLY BACKORDERED</div>');
            $('#btn-submit-bo, #frm-backorder').hide();
		})
		.fail(function(msg) {
            $('#new-bo-message').html('<div class="alert alert-danger">BACKORDER FAILED</div>');
        })
})

// Close backorder button
$(document).on("click", ".close-bo", function(){
    $('#scan-item').prop('readonly', true);
    $("#scan-item").focus();
    $('#scan-item').prop('readonly', false);
})

// Copy orderdetailid to attribute for alerting error button
$(document).on("click", ".error-button", function(){
    var orderdetailid = $(this).attr('data-orderdetailid');
    $('#btn-submit-error').attr('data-orderdetailid', orderdetailid);

    $('#message-error').html('');
    $('#error_description').val('');
    $('#btn-submit-error, #frm-error').show();
})

// Submit inventory issue
$(document).on("click", "#btn-submit-error", function(){
    var notes = $('#error_description').val();
    var item_issue = $('#item_issue').val();
    var orderdetailid = $(this).attr('data-orderdetailid');

      $.ajax({
		method: "post",
		url: "pulling/set-inventory-issue.asp",
		data: {notes: notes, orderdetailid: orderdetailid, item_issue: item_issue}
		})
		.done(function(msg) {
            $('#message-error').html('<div class="alert alert-success">NOTES SUCCESSFULLY SAVED</div>');
            $('#btn-submit-error, #frm-error').hide();
		})
		.fail(function(msg) {
            $('#message-error').html('<div class="alert alert-danger">SUBMIT FAILED</div>');
        })
})







