<%
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT InvoiceID, jewelry.ProductID, DetailID, TBL_OrderSummary.qty, title, ProductDetail1, ProductDetails.qty AS stock_qty, OrderDetailID, backorder, notes, customer_first, email, ID_Description, BinNumber_Detail, Gauge, Length, ProductDetails.ProductDetailID, BackorderReview, PackagedBy, location, ISNULL(replace(jewelry.type,'None',''),'') + ' ' + ISNULL(jewelry.title,'') + ' ' + ISNULL(ProductDetails.ProductDetail1,'') + ' ' + ISNULL(ProductDetails.Gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') as 'item_description' FROM TBL_OrderSummary INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID INNER JOIN sent_items ON TBL_OrderSummary.InvoiceID = sent_items.ID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE BackorderReview = 'Y' ORDER BY InvoiceID ASC" 

set rsGetBackorderReviews = Server.CreateObject("ADODB.Recordset")
rsGetBackorderReviews.CursorLocation = 3 'adUseClient
rsGetBackorderReviews.Open objCmd
If Not rsGetBackorderReviews.EOF Or Not rsGetBackorderReviews.BOF Then %>

<h5 class="mt-5 mb-2"><%= rsGetBackorderReviews.RecordCount %> backorders to be reviewed</h5>

<div class="table-responsive-sm">
<table class="table table-striped table-hover" id="bo-print">
	<thead class="thead-dark">
		<tr>
			<th class="no-print"><button class="btn btn-sm btn-secondary" id="btn_print">PRINT</button></th>
			<th class="no-print">Invoice</th>
			<th>Qty Ordered</th>
			<th>Current Stock</th>
			<th>Item</th>
			<th>Location</th>
			<th class="no-print">Notes</th>
			<th class="no-print">Packaged By</th>
		</tr>
	</thead>
  <% 
While NOT rsGetBackorderReviews.EOF 
%>        <tr id="row_<%=(rsGetBackorderReviews.Fields.Item("OrderDetailID").Value)%>">
		  <td class="no-print">
			<button class="btn btn-success btn-sm btn-submit-bo mr-3 btn-update-bo-modal" data-toggle="modal" data-target="#modal-submit-backorder" data-itemid="<%=(rsGetBackorderReviews.Fields.Item("OrderDetailID").Value)%>" data-qty="<%= rsGetBackorderReviews.Fields.Item("stock_qty").Value %>" data-title="<%= Server.HTMLEncode(rsGetBackorderReviews.Fields.Item("item_description").Value) %>"><strong>APPROVE</strong></button><button class="btn btn-danger btn-sm  btn-clear-bo" data-item="<%=(rsGetBackorderReviews.Fields.Item("OrderDetailID").Value)%>" data-invoiceid="<%=(rsGetBackorderReviews.Fields.Item("InvoiceID").Value)%>" data-productdetailid="<%=rsGetBackorderReviews.Fields.Item("ProductDetailID").Value %>" data-agenda="clear">Deny</button>&nbsp;&nbsp;&nbsp;&nbsp;<strong>
				<span id="spinner_<%=(rsGetBackorderReviews.Fields.Item("OrderDetailID").Value)%>" style="display:none"><i class="fa fa-spinner fa-lg fa-spin"></i></span>
			</td>
          <td class="no-print">
			<a class="text-secondary" href="invoice.asp?ID=<%=(rsGetBackorderReviews.Fields.Item("InvoiceID").Value)%>">Invoice # <%=(rsGetBackorderReviews.Fields.Item("InvoiceID").Value)%></a>&nbsp;&nbsp;(<a class="text-secondary" href="order history.asp?var_email=<%=(rsGetBackorderReviews.Fields.Item("email").Value)%>" target="_blank">History</a>)
		  </td>
		  <td>
			<span class="alert alert-success py-1 px-2"><%=(rsGetBackorderReviews.Fields.Item("qty").Value)%></span>
		  </td>
			<td>
				<input class="form-control form-control-sm" style="width: 40px" type="text" id="deny_qty_<%=rsGetBackorderReviews.Fields.Item("ProductDetailID").Value %>" value="<%=(rsGetBackorderReviews.Fields.Item("stock_qty").Value)%>">
			</td>
		  <td>
			<%=(rsGetBackorderReviews.Fields.Item("title").Value)%>&nbsp;<%=(rsGetBackorderReviews.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetBackorderReviews.Fields.Item("Length").Value)%>&nbsp;&nbsp;<%=(rsGetBackorderReviews.Fields.Item("ProductDetail1").Value)%>
		</td>

			<td>
				<%=(rsGetBackorderReviews.Fields.Item("ID_Description").Value)%>&nbsp;<%=(rsGetBackorderReviews.Fields.Item("location").Value)%>&nbsp;
			<% if (rsGetBackorderReviews.Fields.Item("BinNumber_Detail").Value) <> 0 then %>
				(BIN <%=(rsGetBackorderReviews.Fields.Item("BinNumber_Detail").Value)%>)
			<% end if %>
		  </td>
			<td class="no-print">
				<%=(rsGetBackorderReviews.Fields.Item("notes").Value)%>
			</td>
			<td class="no-print">
				<%=(rsGetBackorderReviews.Fields.Item("PackagedBy").Value)%>
			</td>
        </tr>
    <% 
  rsGetBackorderReviews.MoveNext()
Wend
%>      </table>
</div>
 <!-- Process backorder Modal -->
<div class="modal fade" id="modal-submit-backorder" tabindex="-1" role="dialog"  aria-labelledby="modal-submit-backorder" >
	<div class="modal-dialog" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title">Submit Backorder</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div class="modal-body small">
			<div id="new-bo-message"></div>
			<!--#include virtual="/admin/invoices/inc-submit-backorder.asp"-->
		</div>
		<div class="modal-footer">
			<button type="button" class="btn btn-primary" id="btn-submit-bo" data-itemid="">Submit</button>
		  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
		</div>
	  </div>
	</div>
</div>
<!-- End Process backorder Modal -->   

	<% End If ' end Not rsGetBackorderReviews.EOF Or NOT rsGetBackorderReviews.BOF 
rsGetBackorderReviews.Close()
Set rsGetBackorderReviews = Nothing
	%>