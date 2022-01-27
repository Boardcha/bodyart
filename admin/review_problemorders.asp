<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn

objCmd.CommandText = "SELECT TOP (100) PERCENT dbo.sent_items.ID, dbo.sent_items.shipped, dbo.sent_items.customer_first, dbo.sent_items.customer_last, dbo.sent_items.email, dbo.sent_items.country, dbo.sent_items.PackagedBy, dbo.TBL_OrderSummary.ErrorReportDate, dbo.TBL_OrderSummary.ErrorDescription, dbo.TBL_OrderSummary.ErrorOnReview, dbo.sent_items.ship_code, dbo.TBL_OrderSummary.qty, dbo.TBL_OrderSummary.item_price, dbo.TBL_OrderSummary.notes, dbo.ProductDetails.ProductDetail1, dbo.ProductDetails.location, dbo.ProductDetails.Gauge, dbo.ProductDetails.Length, dbo.jewelry.title, dbo.ProductDetails.ProductDetailID, dbo.ProductDetails.BinNumber_Detail, dbo.ProductDetails.wlsl_price, dbo.TBL_OrderSummary.OrderDetailID, dbo.TBL_OrderSummary.ProductID, dbo.TBL_OrderSummary.item_problem, dbo.TBL_OrderSummary.ErrorQtyMissing, dbo.TBL_Barcodes_SortOrder.ID_Description FROM dbo.sent_items INNER JOIN dbo.TBL_OrderSummary ON dbo.sent_items.ID = dbo.TBL_OrderSummary.InvoiceID INNER JOIN dbo.ProductDetails ON dbo.TBL_OrderSummary.DetailID = dbo.ProductDetails.ProductDetailID INNER JOIN dbo.jewelry ON dbo.TBL_OrderSummary.ProductID = dbo.jewelry.ProductID INNER JOIN dbo.TBL_Barcodes_SortOrder ON dbo.ProductDetails.DetailCode = dbo.TBL_Barcodes_SortOrder.ID_Number WHERE (dbo.TBL_OrderSummary.ErrorOnReview = 1) ORDER BY dbo.sent_items.ID"
set rsGetRecords = objCmd.Execute()

set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.Open objCmd
%>

<html>
<head>
<title>Review PROBLEM orders</title>
<script type="text/javascript" src="../js/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="../js/bootstrap-v4.min.js"></script>
</head>
<body>

<!--#include file="admin_header.asp"-->
<div class="mx-2">
<% If Session("SubAccess") <> "N" then ' DISPLAY ONLY TO PEOPLE WHO HAVE ACCESS TO THIS PAGE %>

<% If NOT rsGetRecords.EOF Then %>
<h5 class="mt-3 mb-2"><%= rsGetRecords.RecordCount %> Problem orders</h5>
<table class="table table-striped table-hover">
	<thead class="thead-dark">
		<tr>
			<th width="15%">Invoice</th>
			<th>Problem</th>
			<th>Qty Ordered</th>
			<th width="20%">Problem item(s) on order</th>
			<th>Location</th>
			<th width="40%">Problem comments</th>
		</tr>
	</thead>
            <% 
While NOT rsGetRecords.EOF 
%>
        
	<tr>
		<td>
			<button class="btn btn-sm btn-secondary btn-update-reship-modal" data-toggle="modal" data-target="#modal-reship-items" data-invoiceid="<%= rsGetRecords.Fields.Item("ID").Value %>">Reship items</button>
			<br/>
			<strong><a href="invoice.asp?ID=<%= rsGetRecords.Fields.Item("ID").Value %>" class="text-secondary"><%= rsGetRecords.Fields.Item("ID").Value %></strong></a>
			<br>
			<%=(rsGetRecords.Fields.Item("country").Value)%><p><a class="text-secondary" href="order history.asp?var_first=<%=(rsGetRecords.Fields.Item("customer_first").Value)%>&var_last=<%=(rsGetRecords.Fields.Item("customer_last").Value)%>">View history</a>
		</td>
		<td>
			<% = (rsGetRecords.Fields.Item("item_problem").Value) %> (Qty: <% = (rsGetRecords.Fields.Item("ErrorQtyMissing").Value) %>)
		</td>
		<td>
			<%=(rsGetRecords.Fields.Item("qty").Value)%>
		</td>
		<td>
            <a class="text-secondary" href="product-edit.asp?ProductID=<%=(rsGetRecords.Fields.Item("ProductID").Value)%>&info=less"><%=(rsGetRecords.Fields.Item("title").Value)%></a>&nbsp; <%=(rsGetRecords.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetRecords.Fields.Item("Length").Value)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Wholesale: <%= FormatCurrency(rsGetRecords.Fields.Item("wlsl_price").Value * (rsGetRecords.Fields.Item("qty").Value), -1, -2, -2, -2) %></td>
			<td>
				<%=(rsGetRecords.Fields.Item("ID_Description").Value)%>&nbsp;<%=(rsGetRecords.Fields.Item("location").Value)%>&nbsp;
			<% if (rsGetRecords.Fields.Item("BinNumber_Detail").Value) <> 0 then %>
				(BIN <%=(rsGetRecords.Fields.Item("BinNumber_Detail").Value)%>)
			<% end if %>
			</td>

              <td><span class="text-danger"><%=(rsGetRecords.Fields.Item("ErrorDescription").Value)%></span></td>
              </tr>
            <% 

  rsGetRecords.MoveNext()
  
Wend
%>
          </table>
     
        <% else ' if there are no records to review %>
		<div class="section-headers">
			No orders are available for review
		</div>
        <% End If ' end rsGetRecords.EOF And rsGetRecords.BOF %>


<% else ' unathorized access error %>
Not accessible
<% end if ' END ACCESS TO PAGE FOR ONLY USERS WHO SHOULD BE ABLE TO SEE IT %>

<!-- Modal to load items from invoice to approve/deny -->
<div class="modal fade" id="modal-reship-items" tabindex="-1" role="dialog"  aria-labelledby="modal-reship-items" >
	<div class="modal-dialog mw-100 w-75" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title">Reship items</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div class="modal-body">
			<div id="message-reship-status"></div>
			<div id="load-reship-items"></div>
		</div>
		<div class="modal-footer">
			<div class="d-inline-block text-left w-50">
				<button type="button" class="btn btn-sm btn-danger btn-reship-items" data-agenda="deny" id="btn-reship-deny" data-invoiceid="">Deny reship</button>
			</div>
			<div class="d-inline-block text-right w-50">
				<button type="button" class="btn btn-sm btn-primary btn-reship-items" data-agenda="approve" id="btn-reship-approve" data-invoiceid="">Reship items</button>
			</div>
		</div>
	  </div>
	</div>
</div>
<!-- End Modal to load items from invoice to approve/deny -->

</div>
<!--#include file="includes/inc_scripts.asp"-->
<script type="text/javascript" src="scripts/generic_auto_update_fields.js"></script>
<script type="text/javascript" src="scripts/review-backorders.js?v=051720"></script>
<script type="text/JavaScript" src="/js/jQuery.print.min.js"></script>
<script type="text/javascript">
	//url to to do auto updating
	var auto_url = "invoices-cs-modules/ajax-reship-update-qty.asp"
	auto_update(); // run function to update fields when tabbing out of them


	// Change reship order modal
	$(document).on("click", ".btn-update-reship-modal", function(event){
		var invoiceid = $(this).attr("data-invoiceid");
		$('.btn-reship-items').attr("data-invoiceid", invoiceid);

		$('#btn-reship-approve').html('Reship items');
		$('#btn-reship-deny').html('Deny reship');
		$('.btn-reship-items').show();
		$('#message-reship-status').html('');

		$('#load-reship-items').load("invoices-cs-modules/ajax-item-problems.asp", {invoiceid: invoiceid})
	
	}); // End change reship order modal

	// Approve or deny reships
	$(document).on("click", ".btn-reship-items", function(event){
		var invoiceid = $(this).attr("data-invoiceid");
		var agenda = $(this).attr("data-agenda");

		if(agenda==='approve') {
			$('#btn-reship-approve').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');
			$('#btn-reship-deny').hide();
		} else {
			$('#btn-reship-deny').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');
			$('#btn-reship-approve').hide();
		}
		$('#load-reship-items').html('');

		$.ajax({
		method: "post",
		url: "invoices-cs-modules/ajax-reship-items.asp",
		data: {invoiceid: invoiceid, agenda:agenda}
		})
		.done(function(msg ) {
			if(agenda==='approve') {
				$('#message-reship-status').html('<div class="alert alert-success h6">ITEMS HAVE BEEN SET TO RESHIP IN A NEW ORDER</div>');
			} else {
				$('#message-reship-status').html('<div class="alert alert-success h6">WINDOW CAN BE CLOSED</div>');
			}
			
			$('.btn-reship-items').hide();
		})
		.fail(function(msg) {
			$('#message-reship-status').html('<div class="alert alert-danger h5">ERROR</div>');
			$('.btn-reship-items').hide();
		});
	
	}); // End Approve or deny reships




</script>
</body>
</html>
<%
rsGetRecords.Close()
%>
<%
rsGetUser.Close()
Set rsGetUser = Nothing
%>
