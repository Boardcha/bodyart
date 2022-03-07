<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TOP(20) * FROM TBL_PurchaseOrders where po_internal_bulk_pull = 1 ORDER BY PurchaseOrderID DESC"
Set rsGetPurchaseOrders = objCmd.Execute()
%>
<html>
<head>
<title>Create Bulk Purchase Order</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h5>
    Create Internal Purchase Order
</h5>
<button class="btn btn-sm btn-primary" id="start-order" data-toggle="modal" data-target="#modal-add">
<% if request.Cookies("bulk-po-id") = "" then %>
    Start bulk order
<% else %>
    Continue bulk order
<% end if %>
</button>
<div class="custom-control custom-checkbox">
  <input name="needs-review" id="needs-review" type="checkbox" class="custom-control-input" value="yes">
  <label class="custom-control-label" for="needs-review">To be reviewed by manager</label>
</div>
<table class="table table-striped table-borderless table-hover mt-4">
                  <% 
    While NOT rsGetPurchaseOrders.EOF
    %>
                    <tr>
                      <td width="1%">
                        <%= FormatDateTime(rsGetPurchaseOrders.Fields.Item("DateOrdered").Value,2)%>
                      </td>
                      <td>
                          <a class="btn btn-sm btn-info ml-5" href="inventory/view-print-bulk-order.asp?ID=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>">
                            Print order to pull</a>
                            <% if rsGetPurchaseOrders("po_needs_review") = "True" then %>

                            <a class="btn btn-sm btn-outline-secondary ml-3" href="/admin/inventory/view-inprogress-bulk-order.asp?ID=<%= rsGetPurchaseOrders("PurchaseOrderID") %>">Manager review order</a>
  
                        <% end if %>
                      </td>
                    </tr>
                    <% 
      rsGetPurchaseOrders.MoveNext()
    Wend
    %>
              </table>
</div><!--admin content-->

<!-- BEGIN ADD ITEM MODAL WINDOW -->
<div class="modal fade" id="modal-add" tabindex="-1" role="dialog"  aria-labelledby="modal-add" >
	<div class="modal-dialog mw-100 w-75" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <div class="modal-title">
            <h5>Add items to pull</h5> <button class="btn btn-sm btn-primary" id="finalize-order" type="button">Finalize order to pull</button><a href="/admin/inventory/view-inprogress-bulk-order.asp" class="btn btn-sm btn-secondary ml-2" id="view-order" target="_blank">View order in progress</a>
            <span class="mr-2" id="msg-finalize"></span>
        </div>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div class="modal-body">
            <div class="form-inline">
                <input class="form-control form-control-sm mr-3" type="text" id="product-id" placeholder="Search Product ID">
                <button class="btn btn-sm btn-secondary" type="button" id="btn-search-product">Search</button>
            </div>
            <div id="display-product"></div>
		</div>
        <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
          </div>
	  </div>
	</div>
</div>
<!-- END ADD ITEM MODAL WINDOW -->

</body>
</html>
<script type="text/javascript">
$("#start-order").click(function(){
  if ($('#needs-review').prop("checked")) {
    $('#msg-finalize').show().html('<div class="alert alert-info mt-2 font-weight-bold">Quantities will be deducted after a manager reviews the order</div>')
    } else {
      $('#msg-finalize').show().html('<div class="alert alert-warning mt-2 font-weight-bold">Quantities will be deducted now ... no manager review</div>')
    }
    
});

$("#btn-search-product").click(function(){
    var productid = $('#product-id').val();
    $('#display-product').load('inventory/ajax-bulk-display-items.asp?productid=' + productid); 
    $('#start-order').html("Continue bulk order");  
});

$(document).on("click", ".btn-add-item", function(){
    var detailid = $(this).attr("data-id");
    var qty_on_hand = $(this).attr("data-on-hand-qty");
    var qty = $('#qty_' + detailid).val();
    
    $('.msg-btn-add-' + detailid).show().html('<i class="fa fa-spinner fa-spin"></i>')

    if ($('#needs-review').prop("checked")) {
      var_needs_review = "yes"
    } else {
      var_needs_review = "no"
    }

    $.ajax({
    method: "post",
    url: "inventory/ajax-bulk-add-item.asp",
    data: {detailid: detailid, qty: qty, var_needs_review: var_needs_review}
    })
    .done(function( msg ) {

      if (var_needs_review == "no") {
        $('.msg-btn-add-' + detailid).html("<div class='alert alert-success p-1'>" + qty + " deducted from stock</div>").delay(8000).fadeOut('slow');
        $('#on-hand-' + detailid).html(qty_on_hand - qty);
      } else {
        $('.msg-btn-add-' + detailid).html("<div class='alert alert-success p-1'>" + qty + " added to order</div>").delay(8000).fadeOut('slow');
      }
        
        
    })
    .fail(function(msg) {
        $('.msg-btn-add-' + detailid).html("<div class='alert alert-danger p-1'>Code error</div>").delay(8000).fadeOut('slow');
    });
});

$(document).on("click", "#finalize-order", function(){
    $('#msg-finalize').show().html('<i class="fa fa-spinner fa-spin mr-3"></i>Building order ...')
    if ($('#needs-review').prop("checked")) {
      var_needs_review = "yes"
    } else {
      var_needs_review = "no"
    }

    $.ajax({
    method: "post",
    url: "inventory/ajax-bulk-finalize-order.asp",
    data: {var_needs_review: var_needs_review}
    })
    .done(function( msg ) {
        window.location.href = "inventory-bulk-pull-po.asp";
    })
    .fail(function(msg) {
        alert("CODE ERROR");
    });
    
});
</script>
<%
DataConn.Close()
Set DataConn = Nothing
%>
