<%
'====== SINCE THIS FILE IS IN ROOT DIRECTORY, MAKE SURE THAT USER IS LOGGED IN VIA ADMIN IN ORDER TO ACCESS CODE ON THIS page
if user_name <> "" then
%>
<script type="text/javascript">
// Change submit backorder modal
$(document).on("click", ".btn-update-bo-modal", function(event){
    var orderdetailid = $(this).attr("data-itemid");
    var qty = $(this).attr("data-qty");
    var title = $(this).attr("data-title");
    var reason = $(this).attr("data-reason");
    $("#frm-submit-backorder, #btn-submit-bo").show();
    $('#new-bo-message').html('');
    $('#btn-submit-bo').attr("data-itemid", orderdetailid);
    $('#btn-submit-bo').attr("data-reason", reason);
    $("#radio").prop("checked", true);
    $('#bo-qty').prop('selectedIndex',0);
    $('#bo-qty-instock').html(qty);
    $('#bo-item-title').html(title);

}); // End submit new backorder

// Submit a new backorder
$(document).on("click", "#btn-submit-bo", function(event){
    var orderdetailid = $(this).attr("data-itemid");
    var bo_reason = $(this).attr("data-reason");
    var bo_qty = $("#bo-qty").val();
    
    $("#frm-submit-backorder").hide();
    $('#new-bo-message').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');
    $('#btn-submit-bo').hide();

    $.ajax({
    method: "POST",
    url: "/includes/ajax-backorder-submit.asp",
    data: {orderdetailid:orderdetailid, bo_reason:bo_reason, bo_qty:bo_qty}
    })
    .done(function(msg ) {
        $("#new-bo-message").html('<span class="alert alert-success">Backorder processed</span>').show();
        $('.bo_blue_' + orderdetailid).hide();
        $('.bo_orange_' + orderdetailid).show();
        $('#row_' + orderdetailid).hide();
    })
    .fail(function(msg) {
        alert('FAILED');
        $("#new-bo-message").hide();
        $('#btn-submit-bo').show();
        $("#frm-submit-backorder").delay(1000).fadeIn(1000);
    });
}); // End submit new backorder

// CLear backorder
$(document).on("click", ".btn-clear-bo", function(event){
    var item = $(this).attr("data-item");
    var productdetailid = $(this).attr("data-productdetailid");
    var agenda = $(this).attr("data-agenda");
    var invoice = $(this).attr("data-invoiceid");
    var stock_qty = $('#deny_qty_' + productdetailid).val();
    
    $('#spinner_' + item).show();
    console.log('Qty we counted ' + stock_qty);
    $.ajax({
    method: "POST",
    url: "/includes/ajax-backorder-process.asp",
    data: {item: item, agenda: agenda, invoice: invoice, stock_qty: stock_qty, detailid: productdetailid}
    })
    .done(function(msg ) {
        $('#row_' + item).hide();
    })
    .fail(function(msg) {
        alert('FAILED');
        $('#spinner_' + item).hide();
    });
}); // End CLear backorder
</script>
<%
end if
%>