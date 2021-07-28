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
rsGetRegular_cmd.CommandText = "SELECT jewelry.ProductID, ProductDetails.ProductDetailID, jewelry.title, jewelry.type, jewelry.active AS MainActive, jewelry.picture, ProductDetails.weight, ProductDetails.DateLastPurchased, ProductDetails.qty, ProductDetails.ProductDetail1, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.active AS DetailActive,  ProductDetails.Date_InventoryCount, ProductDetails.Inventory_TimesScanned FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID WHERE ProductDetailID = ?" 
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
		Response.write "<span class=""alert"">Failed scan match</span>"
		Complete = "yes"
	End if
End if

	  Set rsGetPreorders_cmd = Server.CreateObject ("ADODB.Command")
	  rsGetPreorders_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
	  rsGetPreorders_cmd.CommandText = "SELECT sent_items.shipped, TBL_OrderSummary.DetailID, TBL_OrderSummary.qty FROM         sent_items INNER JOIN TBL_OrderSummary ON sent_items.ID = TBL_OrderSummary.InvoiceID WHERE (sent_items.shipped = N'ON ORDER' OR sent_items.shipped = N'PRE-ORDER APPROVED' OR sent_items.shipped = N'PRE-ORDER REVIEW' OR sent_items.shipped = N'ON HOLD') AND DetailID = " & rsGetRegular.Fields.Item("ProductDetailID").Value 
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
<link rel="stylesheet" href="css/scanners.css" type="text/css" />
<style type="text/css">

<!--
body {
		-webkit-text-size-adjust:none;
	  	font-family: Helvetica, Arial, Verdana, sans-serif;
	  	font-size: 15px;
	  	color: black;
	  }
	
.smallfont {
	   font-size: 13px;
	   }	
	   
.accent {
		font-weight: bold;
		font-size: 16px;
			}   

.accepted {
		   color: #090;
		   font-weight: bold;
		   font-size: 20px;
		   }
.alert {
	   color: #CC0000;
	   font-weight: bold;
	   font-size: 20px;
	   }
-->
</style>
<!--#include file="../includes/inc_scripts.asp"-->
<script type="text/javascript">
	$(document).ready(function() {

		// Update the weight field
		$("#weight").change(function() {
			var weightchange = $(this).val();
			var detailid = $(this).attr("data-detailid");
			var field_name = $(this).attr("name");
			
				$.ajax({
				method: "POST",
				url: "includes/inc_update_weight.asp",
				data: {weight: weightchange, detailid: detailid}
				})
				.done(function( msg ) {
				
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

					$("input[name='"+ field_name +"']").addClass("ajax_input_fail");
					setTimeout(function(){}, 3000);	
				
				//	alert( "error" + msg + "Detail-id: " + detailid + " Qty: " + qtychange);
				});
		}); // end weight update
		
		// Update the qty field
		$("#qty").change(function() {
			var qtychange = $(this).val();
			var detailid = $(this).attr("data-detailid");
			var field_name = $(this).attr("name");
			
				$.ajax({
				method: "POST",
				url: "includes/inc_update_qty.asp",
				data: {qty: qtychange, detailid: detailid}
				})
				.done(function( msg ) {
				
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

					$("input[name='"+ field_name +"']").addClass("ajax_input_fail");
					setTimeout(function(){}, 3000);	
				
				//	alert( "error" + msg + "Detail-id: " + detailid + " Qty: " + qtychange);
				});
		}); // end weight update

	});	
</script>
</head>

<body>
<form action="ItemCount.asp" method="post" name="FRM_Scan">
<input name="Item" type="text" id="Item" placeholder="Scan item #" autofocus/>
<button type="submit" style="display: none">></button>
<% if not rsGetRegular.eof then
	var_ProductDetailID = rsGetRegular.Fields.Item("ProductDetailID").Value
	end if
%>
<br/><br/>
<% if not rsGetRegular.eof then %>
	<div class="inventory-count-regular">
		Qty: <input name="qty" type="text" id="qty" data-detailid="<%= var_ProductDetailID %>" size="4" />
		<%' if rsGetRegular.Fields.Item("weight").Value = 0 then 'if weight field is empty %>
			<br/><br/>			
			Weight: <input name="weight" id="weight" type="text" size="4" placeholder="<%= rsGetRegular.Fields.Item("weight").Value %>" data-detailid="<%= var_ProductDetailID %>" />
			
		<%' end if 'if weight field is empty %> 
	</div>
<% end if ' if a record has been found %>
<p>
  <% If Not rsGetRegular.EOF Or Not rsGetRegular.BOF Then %>  
  <% If RecentlySold = "yes" then   ' SHOW BELOW IF ITEM HASN'T BEEN SOLD IN THE LAST 24 HOURS%>
</p>
 <span class="inventory-count-qty"><%= rsGetRegular.Fields.Item("qty").Value %> in stock</span>
	<br/><br/>
<div>
	<div style="float: left; width: 90px;">
		<img src="http://bodyartforms-products.bodyartforms.com/<%=(rsGetRegular.Fields.Item("picture").Value)%>" alt="Image" width="75" height="75">
	</div> 
	<div style="float: left; width: 70%">
		<%=(rsGetRegular.Fields.Item("title").Value)%>&nbsp;<%=(rsGetRegular.Fields.Item("ProductDetail1").Value)%>&nbsp;<%=(rsGetRegular.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetRegular.Fields.Item("Length").Value)%>
	</div>
	  
	  <% If Not rsGetPreorders.EOF Or Not rsGetPreorders.BOF Then %> 
		<span class="accent">(+ <%= preorder_qty %> reserved for pre-orders)</span>
		<% End If ' end Not rsGetPreorders.EOF Or NOT rsGetPreorders.BOF %>

	<% else ' DISPLAY BELOW IF ITEM HAS BEEN SOLD IN LAST 24 HOURS%> 
	<span class="alert">Write down #<%=(rsGetRegular.Fields.Item("ProductDetailID").Value)%> and re-scan later<br>
	  Last sold <%= (rsGetRegular.Fields.Item("DateLastPurchased").Value) %></span>
	</p>
	<% end if %>
<div style="clear: both;"></div>
</div>
<% Else
If Request.form("Item") <> "" then %>
       <span class="alert">No item found</span>
  <%  End if
 %>

<%  
End If ' end rsGetRegular.EOF And rsGetRegular.BOF %>
</form> 
<% if session("hide_invreg_notes") = "" then %>
	<div class="page-notes">
	<strong>How this page works - Inventory count (Regular stock)</strong>&nbsp;&nbsp;&nbsp;&nbsp;<a href="ItemCount.asp?hide=yes">Hide this</a>
	<br/>
	Use with the scanners. Scan the SMALL BARCODE into the top field (just like you would an invoice). The product pulls up and tells you how many we have in stock. If it's correct, scan the next bin on move on. If you need to update the quantity or the weight, enter the value, and then press the done button on the screen. The field WILL TURN GREEN once it's updated.
	</div>
<% end if 'if session isn't set to hide this message %>
<br/>
<br/>
<br/>
<br/>
</body>
</html>
  <%
If Not rsGetRegular.EOF Or Not rsGetRegular.BOF Then
	rsGetPreorders.Close()
	Set rsGetPreorders = Nothing
End if

rsGetRegular.Close()
Set rsGetRegular = Nothing
%>

