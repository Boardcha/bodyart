<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("bin") <> "" then
	var_bin = request.form("bin")
	
	' Add the date scanned to the bin for tracking
	set rsAddDate = Server.CreateObject("ADODB.Command")
	rsAddDate.ActiveConnection = DataConn
	rsAddDate.CommandText = "UPDATE TBL_BinNumbers SET BinCountDate = '" & now() & "' WHERE BinNumberID = " & var_bin 
	rsAddDate.Execute()
	
else
	var_bin = 5000
end if

if request.querystring("section") = "case" then
	var_section = " AND DetailCode = " & request.querystring("number")
else
	' Do not show any cases if scanning limited
	var_section = " AND DetailCode <> 34 AND DetailCode <> 35 AND DetailCode <> 36 AND DetailCode <> 37 "
end if
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP (100) PERCENT jewelry.ProductID, ProductDetails.ProductDetailID, ProductDetails.BinNumber_Detail, jewelry.title, jewelry.picture, ProductDetails.Gauge + N' ' + ProductDetails.Length + N' ' + ProductDetails.ProductDetail1 AS ProductDescription, ProductDetails.qty, ProductDetails.active AS ActiveDetail, jewelry.active AS ActiveMain, ProductDetails.Date_InventoryCount, ProductDetails.DateLastPurchased, jewelry.type FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID WHERE BinNumber_Detail = ? AND (ProductDetails.active = 1) AND (jewelry.active = 1) "  & var_section & " ORDER BY ProductDetails.ProductDetailID ASC"
	objCmd.Parameters.Append(objCmd.CreateParameter("bin",3,1,10,var_bin))
	Set rsGetItems = objCmd.Execute()

%>
<!DOCTYPE html>
<head>
<title>Inventory</title>
<link href="css/scanners.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../js/jquery-2.2.3.min.js"></script>
<script>
$(document).ready(function(){
	
	$("#scan-detail").focus();
	localStorage.setItem('detailid', '');


	//	$(this).addClass("highlight");


	$(".sell-info").mousedown(function() {
		var detailid = $(this).attr("data-id");
		$(".sold-details_" + detailid).load("includes/inc_sold_details.asp?detailid=" + detailid);
	});
	
	$("input[name='detail-research']").change(function() {
		var detailid = $(this).val();
		$(".detail-research").load("includes/inc_research_detailid.asp?detailid=" + detailid);
		$("input[name='detail-research']").val('');
	});

	// Move current <tr> to top of table and hide table row after scanning next detail id (bag)
	$("#scan-detail").change(function() {

		// If a . found in code then split and if not, just pull value from field. This accounts for two styles of tags in bins
		if (document.getElementById('scan-detail').value.includes(".")) {
			scanned_item_array = $('#scan-detail').val().split('.');
        	current_detailid = scanned_item_array[1];
		} else {
			var current_detailid = $("#scan-detail").val();
		}
		var previous_detailid = localStorage.detailid;
		
		$("#" + current_detailid).addClass("highlight");
		$("#" + current_detailid).prependTo(".scanner-table"); // move row to top
		
		$("#" + previous_detailid).hide();
		localStorage.setItem('detailid', $(this).val());
		$("input[name='scan-detail']").val('');	

		// Update the timestamp for the detail ID scan
		$.ajax({
		method: "POST",
		url: "includes/ajax_update_timestamp.asp",
		data: {detailid: current_detailid}
		});
	});
	
	// Update the qty field
	$(".update-qty input").change(function() {
		var qtychange = $(this).val();
		var detailid = $(this).attr("data-id");
		var field_name = $(this).attr("name");
		
			$.ajax({
			method: "POST",
			url: "includes/inc_update_qty.asp",
			data: {qty: qtychange, detailid: detailid}
			})
			.done(function( msg ) {
				$("#scan-detail").focus(); // move auto focus back to detail field
				
				// Highlight field green for success
				$("input[name='"+ field_name +"']").removeClass("ajax_input_fail");
				$("input[name='"+ field_name +"']").addClass("ajax_input_success");

				setTimeout(function(){
					$("input[name='"+ field_name +"']").addClass("ajax_input_fadeout");
					$("input[name='"+ field_name +"']").removeClass("ajax_input_success");}, 3000);					
					$("input[name='"+ field_name +"']").removeClass("ajax_input_fadeout");
					
				//	alert( "success" + msg + "Detail-id: " + detailid + " Qty: " + qtychange);
			})
			.fail(function(msg) {
			// Highlight field red for failure
			
				$("input[name='"+ field_name +"']").addClass("ajax_input_fail");
				setTimeout(function(){}, 3000);	
			
			//	alert( "error" + msg + "Detail-id: " + detailid + " Qty: " + qtychange);
			});
	}); // end qty update
	
    $('form[name=update-details]').submit(function(){
		var detailid = $(":input[name='detailid']").val();
		var qty = $(":input[name='fix-qty']").val();
		var bin = $('form[name=update-details]').attr("data-bin");

			$.ajax({
			method: "POST",
			url: "includes/inc_assign_detail.asp",
			data: {qty: qty, detailid: detailid, bin: bin}
			})
			.done(function( msg ) {
				$(".update-return-text").removeClass("failed-text");
				$(".update-return-text").html("Update successful");
			})
			.fail(function(msg) {
				$(".update-return-text").html("** Update FAILED **");
				$(".update-return-text").addClass("failed-text");
			});  
        return false;
    }); // end form submit
	
});	
</script>
</head>
<body>
<div class="inventory-count-bin">
	<form action="inventory-count-limited-bin.asp?section=<%= request.querystring("section") %>&number=<%= request.querystring("number") %>" method="post">
			<div style="padding:3px;margin-bottom:1.5em">
					<a href="?section=case&number=37" style="margin-right:1.5em">Case 4 (Gold)</a>
					<a href="?section=case&number=36" style="margin-right:1.5em">Case 3</a>
					<a href="?section=case&number=35" style="margin-right:1.5em">Case 2</a>
					<a href="?section=case&number=34" style="margin-right:1.5em">Case 1</a>
					<a href="?">Scan limited</a>
			</div>
	  <input name="bin" type="text" placeholder="Scan limited BIN #">
	  <% if var_bin <> 5000 then %>
		&nbsp;&nbsp;&nbsp; <strong>BIN # <%= var_bin %></strong>
		<% end if %>
	</form>
	<p>
		<input  id="scan-detail" name="scan-detail" type="text" placeholder="Scan Detail ID #">
	</p>
	<br/>
<% If Not rsGetItems.EOF Or Not rsGetItems.BOF Then %>
<table class="scanner-table" style="width: 100%; border-collapse: collapse;">
<thead>
	<th>Qty</th>
	<th>ID</th>
	<th>&nbsp;</th>
	<th>Name</th>
	<th>Sold</th>
</thead>

  <% 

While NOT rsGetItems.EOF

	if rsGetItems.Fields.Item("DateLastPurchased").Value > (now() - 1) then
		recently_sold = "recently-sold"
	else
		recently_sold = ""
	end if
%>
<tbody>
<tr id="<%= rsGetItems.Fields.Item("ProductDetailID").Value %>">
	<td class="update-qty">
		<input type="text" name="qty-change_<%= rsGetItems.Fields.Item("ProductDetailID").Value %>" size="1" value="<%= rsGetItems.Fields.Item("qty").Value %>" data-id="<%= rsGetItems.Fields.Item("ProductDetailID").Value %>">
	</td>
	<td>
		<%= rsGetItems.Fields.Item("ProductDetailID").Value %>
	</td>
	<td>
		<img src="http://bodyartforms-products.bodyartforms.com/<%= rsGetItems.Fields.Item("picture").Value %>" width="50" height="50" />
	</td>
	<td class="title">
		<a href="../admin/product-edit.asp?ProductID=<%= rsGetItems.Fields.Item("ProductID").Value %>&info=less" target="_blank"><%=(rsGetItems.Fields.Item("title").Value)%></a>&nbsp;<%=(rsGetItems.Fields.Item("ProductDescription").Value)%>
	</td>
	<td class="sell-info <%= recently_sold %>" data-id="<%= rsGetItems.Fields.Item("ProductDetailID").Value %>">
		<%= rsGetItems.Fields.Item("DateLastPurchased").Value %>
		<span class="sold-details_<%= rsGetItems.Fields.Item("ProductDetailID").Value %>"></span>
	</td>
</tr> 
</tbody>

<% 
  rsGetItems.MoveNext()
Wend 
%>
</table>
<% End If ' end Not rsGetItems.EOF Or NOT rsGetItems.BOF %>

<% if var_bin <> 5000 then %>
<br/>
<p>
	<form name="update-details" class="box" data-bin="<%= var_bin %>">
		<input name="detailid" type="text" placeholder="Assign Detail ID # to bin" required>
		<input name="fix-qty" type="text" placeholder="Enter in qty" required>
		<button type="submit">Update</button>
		<p>
		<span class="update-return-text success-text"></span>
		</p>
	</form>
</p>
<p>	
	<div class="box">
		<input type="text" name="detail-research" placeholder="Research Detail ID #">
		<div class="detail-research"></div>
	</div>
</p>
<% end if ' only show if a bin has been submitted %>	
</div>
<div class="page-notes">
<strong><u>How to use this page</u></strong>
<ul>
	<li>
		NOTE: Refreshing the page will reset the inventory count and start it over.
	</li>
	<li>
		Scan or type in the bin # you want to inventory. The page will pull up all the items that our site shows to be active and with qty's in stock in that bin.
	</li>
	<li>
		The field should auto set to the detail ID box. Scan the detail ID barcode on an item. It will move to the top of the page and be highlighted in green. Verify the qty (and if needed update it) and then rescan the bag. The item will be removed from the page.
	</li>
	<li>
		If you have extra inventory in the bin that's not shown to be in that bin, you can easily assign it. At the bottom of the page, scan the detail ID bardcode, enter the qty, and hit submit. It'll show it's updated. NOTE: It won't show up at the top of the page but it did work if you get the "Update successful" notice.
	</li>
	<li>
		If you need to find out more information about an item (via it's detail ID) you can use the research field at the bottom and it will pull up all necessary info without having to leave the page.
	</li>	
	<li>
		Scan all the items in the bin and then you'll see what's left on the page that needs to be made inactive (or researched), and also what items need to be made active as well.
	</li>
<ul>
</div>
</body>
</html>
<%
DataConn.Close()
Set rsGetItems = Nothing
%>
