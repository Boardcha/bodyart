<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.querystring("hide") = "yes" then
	session("hide_invreg_notes") = "yes"
end if

complete = "no"
RecentlySold = "no"

If Request.Form("Item") <> "" then
	If Request.Form("OrigScan") = "" Then	
		ItemScan = Request.Form("Item")
	Else
		ItemScan = Request.Form("Item")
	End if
Else
	ItemScan = 0
End if



Set rsGetRegular_cmd = Server.CreateObject ("ADODB.Command")
rsGetRegular_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRegular_cmd.CommandText = "SELECT jewelry.ProductID, ProductDetails.ProductDetailID, jewelry.title, jewelry.type, jewelry.active AS MainActive, jewelry.picture, ProductDetails.weight, ProductDetails.DateLastPurchased, ProductDetails.qty, ProductDetails.ProductDetail1, ProductDetails.Gauge, ProductDetails.Length, largepic, ProductDetails.active AS DetailActive,  ProductDetails.Date_InventoryCount, ProductDetails.Inventory_TimesScanned, DetailCode, location, ID_Description FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE ProductDetailID = ?" 
rsGetRegular_cmd.Prepared = true
rsGetRegular_cmd.Parameters.Append rsGetRegular_cmd.CreateParameter("param1", 5, 1, -1, ItemScan) ' adDouble

Set rsGetRegular = rsGetRegular_cmd.Execute


If Not rsGetRegular.EOF Or Not rsGetRegular.BOF Then

	'if a record is found, then update the timestamp
	set UpdateRow = Server.CreateObject("ADODB.Command")
	UpdateRow.ActiveConnection = MM_bodyartforms_sql_STRING
	UpdateRow.CommandText = "UPDATE ProductDetails SET Date_InventoryCount = '" & now() & "' WHERE ProductDetailID = " & Request.Form("Item") & "" '
	UpdateRow.Execute() 

If (rsGetRegular.Fields.Item("DateLastPurchased").Value) < now() - 2 OR IsNull(rsGetRegular.Fields.Item("DateLastPurchased").Value) then
	
	RecentlySold = "yes"

End if


	'response.write Request.Form("OrigScan")

If Request.Form("OrigScan") <> "" then 

rsGetRegular.ReQuery

Else
	If Request.Form("OrigScan") <> "" Then
		Response.write "<span class=""alert alert-danger"">Failed scan match</span>"
		Complete = "yes"
	End if
End if

	  Set rsGetPreorders_cmd = Server.CreateObject ("ADODB.Command")
	  rsGetPreorders_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
	  rsGetPreorders_cmd.CommandText = "SELECT sent_items.shipped, TBL_OrderSummary.DetailID, TBL_OrderSummary.qty FROM         sent_items INNER JOIN TBL_OrderSummary ON sent_items.ID = TBL_OrderSummary.InvoiceID WHERE (sent_items.shipped = N'ON ORDER' OR sent_items.shipped = N'CUSTOM ORDER IN REVIEW' OR sent_items.shipped = N'ON HOLD') AND DetailID = " & rsGetRegular.Fields.Item("ProductDetailID").Value 
	  rsGetPreorders_cmd.Prepared = true
	   
	  Set rsGetPreorders = rsGetPreorders_cmd.Execute
	  
	  preorder_qty = 0

%>
<!--#include virtual="/emails/function-send-email.asp"-->
<%

While NOT rsGetPreorders.EOF 

	preorder_qty = preorder_qty + rsGetPreorders.Fields.Item("qty").Value

rsGetPreorders.MoveNext()
Wend

	' check to see if item has an issue that needs an email sent about it
If rsGetRegular.Fields.Item("MainActive").Value = 0 OR rsGetRegular.Fields.Item("DetailActive").Value = 0 OR rsGetRegular.Fields.Item("type").Value = "Clearance" OR rsGetRegular.Fields.Item("type").Value = "limited" OR rsGetRegular.Fields.Item("type").Value = "Discontinued" OR rsGetRegular.Fields.Item("type").Value = "One time buy" then
	
	
mailer_type = "inventory-count-notification"
%>
<!--#include virtual="/emails/email_variables.asp"-->
<%

End if ' done checking problem to see if email needs to be sent

End if ' If regular recordset is not empty
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Regular inventory count</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0" />
<meta name="mobile-web-app-capable" content="yes">
<link href="/CSS/baf.min.css?v=092519" rel="stylesheet" type="text/css" />
</head>

<body>
<!--#include file="includes/scanners-header.asp" -->
<div class="p-2">
	<form class="form-group" action="ItemCount.asp" method="post" name="FRM_Scan">
		<input class="form-control" name="Item" type="text" id="Item" placeholder="Scan item #" autofocus/>
	
	<button type="submit" style="display: none">></button>
	</form>
<% if not rsGetRegular.eof then
var_ProductDetailID = rsGetRegular.Fields.Item("ProductDetailID").Value
%>
	<div class="row">
		<div class="col mb-4">
			<img class="float-left mr-3" style="width:150px;height:150px" src="http://bodyartforms-products.bodyartforms.com/<%=(rsGetRegular.Fields.Item("largepic").Value)%>" alt="Image">
			<%=(rsGetRegular.Fields.Item("title").Value)%>&nbsp;<%=(rsGetRegular.Fields.Item("ProductDetail1").Value)%>&nbsp;<%=(rsGetRegular.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetRegular.Fields.Item("Length").Value)%>
			<br>
			<span class="badge badge-info"><%= rsGetRegular("ID_Description")%>&nbsp;&nbsp;<%= rsGetRegular("location") %></span>
			<div class="form-inline my-2">	
				<input class="form-control" name="qty" type="text" id="qty" data-detailid="<%= var_ProductDetailID %>" placeholder="Enter new quantity" />
			</div>
			<div class="font-weight-bold d-inline-block" id="qty-message"></div>

			<% If RecentlySold = "yes" then   ' SHOW BELOW IF ITEM HASN'T BEEN SOLD IN THE LAST 24 HOURS%>
			<div class="alert alert-success d-inline-block font-weight-bold">
				<span id="new-qty-number"><%= rsGetRegular.Fields.Item("qty").Value %></span> in stock 
				<% If Not rsGetPreorders.EOF Or Not rsGetPreorders.BOF Then %> 
					+ <%= preorder_qty %> reserved for custom orders
				<% End If ' end Not rsGetPreorders.EOF Or NOT rsGetPreorders.BOF %>
		   </div>
	   
		   <% else ' DISPLAY BELOW IF ITEM HAS BEEN SOLD IN LAST 24 HOURS%> 
		   <div>
			   <span class="alert alert-danger d-inline-block font-weight-bold">Write down #<%=(rsGetRegular.Fields.Item("ProductDetailID").Value)%> and re-scan later<br>
			   Last sold <%= (rsGetRegular.Fields.Item("DateLastPurchased").Value) %></span>
			   </p>
		   </div>
		   <% end if %>

			<div>
		   		<span class="alert alert-warning py-0 px-1 font-weight-bold text-dark ml-3 error-button" style="color:#BDBDBD" data-toggle="modal" data-target="#modal-submit-error"  data-detailid="<%= rsGetRegular("ProductDetailID") %>">Report issue</span>
			</div>
		</div>
	</div>
 
<% Else
If Request.form("Item") <> "" then %>
       <span class="alert alert-danger">No item found</span>
  <%  End if %>

<%  
End If ' end rsGetRegular.EOF  %>

<% if session("hide_invreg_notes") = "" then %>
	<div class="alert alert-info mt-4">
	<strong>How this page works - Inventory count (Regular stock)</strong>&nbsp;&nbsp;&nbsp;&nbsp;<a href="ItemCount.asp?hide=yes">Hide this</a>
	<br/>
	Scan the <strong>SMALL BARCODE</strong> on the label. The product pulls up and tells you how many we have in stock. If it's correct, scan the next bin on move on. If you need to update the quantity or the weight, enter the value, and then press the done button on the screen. The field WILL TURN GREEN once it's updated.
	</div>
<% end if 'if session isn't set to hide this message %>

</div><!-- end body padding -->

<!-- Process Error Alert Modal -->
<div class="modal fade" id="modal-submit-error" tabindex="-1" role="dialog"  aria-labelledby="modal-submit-error" >
	<div class="modal-dialog" role="document">
	  <div class="modal-content">
		<div class="modal-body">
			<div id="message-error"></div>
				<!--#include file="includes/form-report-error.inc" -->
		</div>
		<div class="modal-footer">
			<button type="button" class="btn btn-primary" id="btn-submit-error" data-detailid="">Submit</button>
		  <button type="button" class="btn btn-secondary close-bo" data-dismiss="modal">Close</button>
		</div>
	  </div>
	</div>
  </div>
  <!-- End Process Error Alert Modal -->

</body>
</html>
<script src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="../js/bootstrap-v4.min.js"></script>
<script type="text/javascript">
	function ResetItemField() {
    $("#Item").val('');
    $('#Item').prop('readonly', true);
    $("#Item").focus();
    $('#Item').prop('readonly', false);
	};

	// Update the qty field
	$("#qty").change(function() {
		var qtychange = $(this).val();
		var detailid = $(this).attr("data-detailid");
		var field_name = $(this).attr("name");

		$('#qty-message').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');

		// add a fail safe in case someone scans a bin into the qty field making the qty count crazy high
		if (qtychange > 300){
			$('#qty-message').html('<span class="alert alert-danger">Will not allow qty over 300</span>');
		}

		else {
		
			$.ajax({
			method: "POST",
			url: "includes/inc_update_qty.asp",
			data: {qty: qtychange, detailid: detailid}
			})
			.done(function( msg ) {
		
				$('#qty-message').html('<span class="alert alert-success">Quantity updated</span>');
				$('#new-qty-number').html(qtychange);
				
				ResetItemField()
				//	alert( "success" + msg + "Detail-id: " + detailid + " Qty: " + qtychange);
			})
			.fail(function(msg) {

				$('#qty-message').html('<span class="alert alert-danger">Update failed. Code or input issue.</span>');
			
			//	alert( "error" + msg + "Detail-id: " + detailid + " Qty: " + qtychange);
			});
		} // if qty is not over 300
	}); // end qty update




	// Copy orderdetailid to attribute for alerting error button
	$(document).on("click", ".error-button", function(){
		var detailid = $(this).attr('data-detailid');
		$('#btn-submit-error').attr('data-detailid', detailid);

		$('#message-error').html('');
		$('#error_description').val('');
		$('#btn-submit-error, #frm-error').show();
	})

	// Submit inventory issue
	$(document).on("click", "#btn-submit-error", function(){
		var notes = $('#error_description').val();
		var item_issue = $('#item_issue').val();
		var detailid = $(this).attr('data-detailid');

		$.ajax({
			method: "post",
			url: "pulling/set-inventory-issue.asp",
			data: {notes: notes, detailid: detailid, item_issue: item_issue, report_source:'Scanner - Inventory count'}
			})
			.done(function(msg) {
				$('#message-error').html('<div class="alert alert-success">NOTES SUCCESSFULLY SAVED</div>');
				$('#btn-submit-error, #frm-error').hide();
			})
			.fail(function(msg) {
				$('#message-error').html('<div class="alert alert-danger">SUBMIT FAILED</div>');
			})
	})
</script>
  <%
If Not rsGetRegular.EOF Or Not rsGetRegular.BOF Then
	rsGetPreorders.Close()
	Set rsGetPreorders = Nothing
End if

rsGetRegular.Close()
Set rsGetRegular = Nothing
%>

