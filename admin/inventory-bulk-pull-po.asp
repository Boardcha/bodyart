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
<table class="table table-striped table-borderless table-hover mt-4">
                  <% 
    While NOT rsGetPurchaseOrders.EOF
    %>
                    <tr>
                      <td class="align-middle">
                          <%= FormatDateTime(rsGetPurchaseOrders.Fields.Item("DateOrdered").Value,2)%>
                          <a class="ml-5" href="inventory/view-print-bulk-order.asp?ID=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>">
                            Print order to pull</a>
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
		  <h5 class="modal-title">Add items to pull <button class="btn btn-sm btn-primary ml-5" id="finalize-order" type="button">Finalize order to pull</button><span class="mr-2" id="msg-finalize"></span></h5>
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

    $.ajax({
    method: "post",
    url: "inventory/ajax-bulk-add-item.asp",
    data: {detailid: detailid, qty: qty}
    })
    .done(function( msg ) {
        $('.msg-btn-add-' + detailid).html("<div class='alert alert-success p-1'>" + qty + " deducted from stock</div>").delay(8000).fadeOut('slow');
        $('#on-hand-' + detailid).html(qty_on_hand - qty);
        
    })
    .fail(function(msg) {
        $('.msg-btn-add-' + detailid).html("<div class='alert alert-danger p-1'>Code error</div>").delay(8000).fadeOut('slow');
    });
});

$(document).on("click", "#finalize-order", function(){
    $('#msg-finalize').show().html('<i class="fa fa-spinner fa-spin mr-3"></i>Building order ...')

    $.ajax({
    method: "post",
    url: "inventory/ajax-bulk-finalize-order.asp",
    data: {}
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
