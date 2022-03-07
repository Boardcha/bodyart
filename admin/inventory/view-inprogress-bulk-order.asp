<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.querystring("ID") <> "" then
  '==== Only is used if the manager needs to review the order after it's been placed
  var_po_id = request.querystring("ID")
  sql = "tbl_po_details.po_orderid"
else
  var_po_id = request.Cookies("bulk-po-id")
  sql = "tbl_po_details.po_temp_id"
end if

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT (jewelry.title + ' ' + ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '')) as description, picture, jewelry.ProductID, tbl_po_details.po_qty, tbl_po_details.po_qty_vendor, ProductDetails.detail_code, ProductDetails.ProductDetailID FROM dbo.jewelry INNER JOIN dbo.ProductDetails ON dbo.jewelry.ProductID = dbo.ProductDetails.ProductID INNER JOIN dbo.tbl_po_details ON dbo.ProductDetails.ProductDetailID = dbo.tbl_po_details.po_detailid  WHERE " & sql & " = ? AND tbl_po_details.po_qty > 0 ORDER BY 'description' ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("po_new_id",3,1,10, var_po_id ))
set rsGetItems = objCmd.Execute()
%>
<html>
<head>
    <title>Bulk order</title>
</head>
<body>
<!--#include file="../admin_header.asp"-->
<div class="p-2">
<h4>Review order</h4>
<% if request.querystring("ID") <> "" then %>
<button class="btn btn-sm btn-primary mr-3" id="approve-order" type="button" data-po_id="<%= request.querystring("ID") %>">Approve inventory deduction</button><span class="mr-2" id="msg-finalize"></span>
<% end if %>
<table class="table table-striped table-borderless table-hover mt-4">
<thead>
  <tr>
    <th></th>
    <th class="text-center">Qty</th>
	<th>Item</th>
  </tr>
</thead>
<%
While NOT rsGetItems.EOF 
%>
  <tr>
    <td class="text-center" width="1%">
      <a href="../productdetails.asp?ProductID=<%=(rsGetItems.Fields.Item("ProductID").Value)%>" target="_blank"><img id="main_img" src="http://bodyartforms-products.bodyartforms.com/<%=(rsGetItems.Fields.Item("picture").Value)%>" width="90" height="90"> </a>
	
	</td>
    <td class="text-center" width="1%">
		<%= rsGetItems("po_qty") %>
    </td>
    <td class="text-left">
      <%= rsGetItems.Fields.Item("description").Value %>
    </td>	
  </tr>
<%
  rsGetItems.MoveNext()
Wend
%>
</table>
</div>
</body>
</html>
<script>
  $(document).on("click", "#approve-order", function(){
    $('#msg-finalize').show().html('<i class="fa fa-spinner fa-spin mr-3"></i>Deducting inventory ...')
    var po_id = $(this).attr("data-po_id");

    $.ajax({
    method: "post",
    url: "ajax-bulk-deduct-all-inventory.asp",
    data: {po_id: po_id}
    })
    .done(function( msg ) {
      $('#msg-finalize').show().html('<br><div class="alert alert-success">INVENTORY HAS BEEN DEDUCTED. Order can now be pulled. <a href="/admin/inventory-bulk-pull-po.asp"><br>Go back to bulk orders</a></div>')
      $("#approve-order").hide();
    })
    .fail(function(msg) {
        alert("CODE ERROR");
    });
    
  });
</script>
<%
DataConn.Close()
%>
