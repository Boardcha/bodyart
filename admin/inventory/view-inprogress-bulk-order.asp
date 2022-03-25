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
objCmd.CommandText = "SELECT (jewelry.title + ' ' + ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '')) as description, amt_waiting, sales as total_sales, last_purchase_received, picture, jewelry.ProductID, tbl_po_details.po_qty, tbl_po_details.po_qty_vendor, ProductDetails.detail_code, ProductDetails.ProductDetailID, wlsl_price, qty, DateLastPurchased FROM dbo.jewelry INNER JOIN dbo.ProductDetails ON dbo.jewelry.ProductID = dbo.ProductDetails.ProductID INNER JOIN dbo.tbl_po_details ON dbo.ProductDetails.ProductDetailID = dbo.tbl_po_details.po_detailid  LEFT JOIN TBL_Sales_From_Last_Restock Sales ON Sales.ProductDetailID = ProductDetails.ProductDetailID LEFT OUTER JOIN vw_po_waiting ON vw_po_waiting.DetailID = ProductDetails.ProductDetailID WHERE " & sql & " = ? AND tbl_po_details.po_qty > 0 ORDER BY 'description' ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("po_new_id",3,1,10, var_po_id ))
set rsGetItems = objCmd.Execute()
%>
<html>
<head>
    <title>Bulk order</title>
	<script type="text/javascript" src="/js/popper.min.js"></script>
</head>
<body>
<!--#include file="../admin_header.asp"-->
<div class="p-2">
<h4>Review order</h4>
<% if request.querystring("ID") <> "" then %>
<button class="btn btn-sm btn-primary mr-3" id="approve-order" type="button" data-po_id="<%= request.querystring("ID") %>">Approve inventory deduction</button><span class="mr-2" id="msg-finalize"></span>
<% end if %>

<table class="table table-sm table-striped table-borderless table-hover mt-4">
<thead class="thead-dark sticky-top">
  <tr>
    <th class="sticky-top"></th>
	<th class="sticky-top"></th>
	<th class="sticky-top"></th>
    <th class="sticky-top text-center">Ordered qty</th>
    <th class="sticky-top text-center">Qty on hand</th>
	<th class="sticky-top">Item</th>
	<th class="sticky-top">Wholesale</th>
	<th class="sticky-top">Sale dates</th>
  </tr>
</thead>
<%
row_id = 1
While NOT rsGetItems.EOF 
	isPOamountCalculatedBasedOn_po_date_received = false
	'If there is a po_date_received for the item, use it in calculation 
	if  Not ISNULL(rsGetItems("total_sales")) AND Not ISNULL(rsGetItems("last_purchase_received")) then
		isPOamountCalculatedBasedOn_po_date_received = true
	end if
%>
  <tbody id="tbody_<%= rsGetItems.Fields.Item("ProductDetailID").Value %>">	
  <tr>
	<td class="align-middle">
	<span class="btn btn-sm btn-secondary mr-2 toggle-product-detail" id="<%= row_id %>" data-detailID="<%= rsGetItems.Fields.Item("ProductDetailID").Value %>"><i class="fa fa-angle-down fa-lg product-detail-expand<%= row_id %>"></i><i class="fa fa-angle-up fa-lg product-detail-expand<%= row_id %>" style="display:none"></i></span>
	</td>  
    <td class="text-center" width="1%">
      <a href="/admin/product-edit.asp?ProductID=<%=(rsGetItems.Fields.Item("ProductID").Value)%>" target="_blank"><img id="main_img" src="http://bodyartforms-products.bodyartforms.com/<%=(rsGetItems.Fields.Item("picture").Value)%>" width="90" height="90"> </a>
	
	</td>
	<td class="text-center align-middle">
	<%If isPOamountCalculatedBasedOn_po_date_received Then %>
		<% If Not IsNULL(rsGetItems("amt_waiting")) Then
			var_amt_waiting = rsGetItems("amt_waiting")
		Else
			var_amt_waiting = 0
		End If%>
		<span data-toggle="tooltip"  data-html="true"
			title="<% Response.Write "Last received: " & rsGetItems("last_purchase_received") & "<br>" & _
			"Sales: " & rsGetItems("total_sales") & "<br>" & _
			"On hand: " & rsGetItems.Fields.Item("qty").Value  & "<br>" & _
			"In Waiting List: " & var_amt_waiting %>" 
			class="fa fa-information d-inline-block" style="font-size:22px;vertical-align:middle;">
		</span>
	<%End If%>
	<%If rsGetItems.Fields.Item("amt_waiting").Value > 0 Then %>
		<a class="badge badge-info p-2" href="waitinglist_view.asp?DetailID=<%= rsGetItems.Fields.Item("ProductDetailID").Value %>" target="_blank"><span class="fa fa-user" aria-hidden="true"><sup class="pl-1 font-weight-bold"><%= rsGetDetail.Fields.Item("amt_waiting").Value %></sup></span></a>
	<%End If%>	
	</td>	
    <td class="text-center align-middle" width="10%">
		<span class="alert alert-secondary font-weight-bold py-0 px-1"><%= rsGetItems("po_qty") %></span>
    </td>
    <td class="text-center align-middle" width="10%">
      <span class="alert alert-success font-weight-bold py-0 px-1"><%= rsGetItems("qty") %></span>
    </td>	
    <td class="text-left align-middle">
      <%= rsGetItems.Fields.Item("description").Value %>
    </td>
    <td class="text-center align-middle" width="1%">
      $<%= rsGetItems("wlsl_price") %>
    </td>
    <td class="align-middle">
      <% if rsGetItems.Fields.Item("DateLastPurchased").Value <> "" then %>
      <span role="button" class="date_expand" id="last_sold_<%= rsGetItems.Fields.Item("ProductDetailID").Value %>" data-container="body" data-toggle="popover" data-placement="left" data-html="true" data-trigger="focus" data-content='Loading <i class="fa fa-spinner fa-spin ml-3"></i>' data-detailid="<%= rsGetItems.Fields.Item("ProductDetailID").Value %>">
        <%= FormatDateTime(rsGetItems.Fields.Item("DateLastPurchased").Value,2)%>
      </span>
    <% end if %>
    </td>
  </tr>
  </tbody>
  <tbody class="tbody-nohover">
	<tr class="td-expand<%= row_id %> bg-white" style="display:none">
		<td colspan="14" class="load<%= row_id %>">
		</td>
	</tr>
</tbody>
<%
  row_id = row_id + 1
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
<!--#include file="inc-product-sales-line-graph.inc" -->
<!--#include file="inc-last-sold-dates-popover.inc" -->
<%
DataConn.Close()
%>
