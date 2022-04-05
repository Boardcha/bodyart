<%
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT InvoiceID, jewelry.ProductID, DetailID, TBL_OrderSummary.qty, title, ProductDetail1, ProductDetails.qty AS stock_qty, OrderDetailID, backorder, notes, customer_first, email, ID_Description, BinNumber_Detail, Gauge, Length, picture, ProductDetails.ProductDetailID, BackorderReview, reason_for_backorder, PackagedBy, location, ISNULL(replace(jewelry.type,'None',''),'') + ' ' + ISNULL(jewelry.title,'') + ' ' + ISNULL(ProductDetails.ProductDetail1,'') + ' ' + ISNULL(ProductDetails.Gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') as 'item_description' FROM TBL_OrderSummary INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID INNER JOIN sent_items ON TBL_OrderSummary.InvoiceID = sent_items.ID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE BackorderReview = 'Y' ORDER BY ProductDetails.BinNumber_Detail ASC, TBL_Barcodes_SortOrder.ID_SortOrder, ID_Description ASC, location ASC, ProductDetails.ProductDetailID" 

set rsGetBackorderReviews = Server.CreateObject("ADODB.Recordset")
rsGetBackorderReviews.CursorLocation = 3 'adUseClient
rsGetBackorderReviews.Open objCmd
If Not rsGetBackorderReviews.EOF Or Not rsGetBackorderReviews.BOF Then %>

<h5 class="mt-3 mb-2"><%= rsGetBackorderReviews.RecordCount %> backorders to be reviewed</h5>

<div class="container-fluid">
	<div class="row no-gutters">
		<% 
While NOT rsGetBackorderReviews.EOF 
%> 
		<div class="col-md-3 col-sm-12 p-1" id="row_<%= rsGetBackorderReviews("OrderDetailID") %>">
			<div class="card  bg-light h-100">
				<div class="card-header">
					<div class="d-none d-md-block">
						<a class="text-secondary" href="/admin/invoice.asp?ID=<%=(rsGetBackorderReviews.Fields.Item("InvoiceID").Value)%>">Invoice # <%=(rsGetBackorderReviews.Fields.Item("InvoiceID").Value)%></a>&nbsp;&nbsp;(<a class="text-secondary" href="/admin/order history.asp?var_email=<%=(rsGetBackorderReviews.Fields.Item("email").Value)%>" target="_blank">History</a>)
					</div>
					<div class="h5">
						<%=(rsGetBackorderReviews.Fields.Item("ID_Description").Value)%>&nbsp;<%=(rsGetBackorderReviews.Fields.Item("location").Value)%>&nbsp;
						<% if (rsGetBackorderReviews.Fields.Item("BinNumber_Detail").Value) <> 0 then %>
							(BIN <%=(rsGetBackorderReviews.Fields.Item("BinNumber_Detail").Value)%>)
						<% end if %>
					</div>
					<button class="btn btn-success btn-sm btn-submit-bo mr-3 btn-update-bo-modal" data-toggle="modal" data-target="#modal-submit-backorder" data-itemid="<%=(rsGetBackorderReviews.Fields.Item("OrderDetailID").Value)%>" data-qty="<%= rsGetBackorderReviews.Fields.Item("stock_qty").Value %>" data-title="<%= Server.HTMLEncode(rsGetBackorderReviews.Fields.Item("item_description").Value) %>" data-reason="<%= rsGetBackorderReviews("reason_for_backorder") %>"><strong>APPROVE <span class="mx-1 px-2 bg-light text-dark"><%= rsGetBackorderReviews("qty") %></span> for BO</strong></button><button class="btn btn-danger btn-sm  btn-clear-bo" data-item="<%=(rsGetBackorderReviews.Fields.Item("OrderDetailID").Value)%>" data-invoiceid="<%=(rsGetBackorderReviews.Fields.Item("InvoiceID").Value)%>" data-productdetailid="<%=rsGetBackorderReviews.Fields.Item("ProductDetailID").Value %>" data-agenda="clear">Deny</button>
						<span class="ml-3" id="spinner_<%=(rsGetBackorderReviews.Fields.Item("OrderDetailID").Value)%>" style="display:none"><i class="fa fa-spinner fa-lg fa-spin"></i></span>
				</div>
				<div class="card-body">
					<div class=" clearfix">
						<img class="float-left mr-2 mb-2" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetBackorderReviews("picture") %>" style="width:70px;height:auto">
				
						<span class="d-block font-weight-bold">Current stock</span> <input class="form-control" style="width: 100px" type="text" id="deny_qty_<%=rsGetBackorderReviews.Fields.Item("ProductDetailID").Value %>" value="<%=(rsGetBackorderReviews.Fields.Item("stock_qty").Value)%>">
					</div>
					<div class="small">
						<strong>Item:</strong> <%=(rsGetBackorderReviews.Fields.Item("title").Value)%>&nbsp;<%=(rsGetBackorderReviews.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetBackorderReviews.Fields.Item("Length").Value)%>&nbsp;&nbsp;<%=(rsGetBackorderReviews.Fields.Item("ProductDetail1").Value)%><br>
						<strong>Reason:</strong> <%= rsGetBackorderReviews("reason_for_backorder")%><br>

						<%= rsGetBackorderReviews("notes") %>
						<strong>Packer:</strong> <%= rsGetBackorderReviews("PackagedBy") %>
					</div>
				</div><!-- card body -->
				</div><!-- card -->
		</div><!-- column -->
		<% 
rsGetBackorderReviews.MoveNext()
Wend
%> 
	</div><!-- row -->
</div><!-- container -->

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
			<% hide_reasons = true %>
			<!--#include virtual="/admin/invoices/inc-submit-backorder.asp"-->
		</div>
		<div class="modal-footer">
			<button type="button" class="btn btn-primary" id="btn-submit-bo" data-itemid="" data-reason="">Submit</button>
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