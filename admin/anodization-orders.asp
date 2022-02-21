<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRecords.Source = "SELECT DISTINCT ID, customer_first, customer_last, date_order_placed, convert(nvarchar(max), our_notes) as our_notes " & _
	"FROM TBL_OrderSummary ORS LEFT JOIN sent_items SNT ON SNT.ID = ORS.InvoiceID AND item_price > 0 AND anodized_completed = 0 AND anodization_id_ordered > 0 WHERE SNT.anodize = 1 ORDER BY date_order_placed ASC"
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.LockType = 1 'Read-only records
rsGetRecords.Open()
rsGetRecords_numRows = 0
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Orders needing anodization</title>
<style>
	.row:hover select  {background-color:#6FA59A}
	.row:hover input[type="checkbox"]{outline:2px solid #6FA59A;outline-offset: -2px;}
</style>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">

<h4>Orders needing anodization</h4>

<div class="card my-3">
	<div class="card-header h5">
		Alter barcode labels
	</div>
	<div class="card-body">
			
			<form action="barcodes_preorders.asp" method="post">
				<div class="form-inline">
					<input class="mr-1"  type="radio" name="DetailSort" value="Equal" id="type_0">=
					<input class="ml-4 mr-1"  type="radio" name="DetailSort" value="Greater" id="type_0">>
					<input class="ml-4 mr-1"  type="radio" name="DetailSort" value="GreaterLess" id="type_2" checked="checked">&lt; &gt; 

					<span class="ml-5" >Invoices:</span>
					<input class="form-control form-control-sm ml-2 w-25" name="Details" type="text" id="Details" placeholder= "Example: 123456, 456789, 789012" /> 
					<span class="mx-3">through</span> 
					<input class="form-control form-control-sm w-25" name="Details2" type="text" id="Details2" >
				</div>
				<div class="mt-2">
					<button class="btn btn-purple" type="submit">Update labels</button>
				</div>
			</form>
	</div>
</div>

<table class="table table-striped table-borderless table-hover mt-3">
  <% 
While NOT rsGetRecords.EOF
%>
  <tr> 
        <td style="width:20%"><%=(rsGetRecords.Fields.Item("customer_first").Value)%> &nbsp;<%=(rsGetRecords.Fields.Item("customer_last").Value)%><br>
        <a  href="invoice.asp?ID=<%= rsGetRecords.Fields.Item("ID").Value %>" target="_blank">Invoice <%=(rsGetRecords.Fields.Item("ID").Value)%></a><br>
        Placed: <%=FormatDateTime((rsGetRecords.Fields.Item("date_order_placed").Value),2)%>
        </td>
        <td>
          <%
Dim rsGetOrderDetails2
Dim rsGetOrderDetails2_numRows

Set rsGetOrderDetails2 = Server.CreateObject("ADODB.Recordset")
With rsGetOrderDetails2
rsGetOrderDetails2.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetOrderDetails2.Source = "SELECT OrderDetailID, qty, title, ProductDetail1, PreOrder_Desc, notes, backorder, ProductID, Gauge, Length, brandname, anodized_completed FROM dbo.QRY_OrderDetails WHERE anodization_id_ordered > 0 AND item_price > 0 AND ID = " & rsGetRecords.Fields.Item("ID").Value & " ORDER BY item_ordered_date ASC"
rsGetOrderDetails2.CursorLocation = 3 'adUseClient
rsGetOrderDetails2.LockType = 1 'Read-only records
rsGetOrderDetails2.Open()
%>
<div class="container">
<%
Do While Not.Eof

	anodized_completed = ""
if rsGetOrderDetails2.Fields.Item("anodized_completed").Value = true then
	anodized_completed = "yes"
end if
%>
	<div class="row h-100 my-2">
		<div class="col my-auto <% if anodized_completed <> "" then %>small<% end if %>">
			<% if anodized_completed = "" then %>
				<input class="mr-2 checkbox_item_id" type="checkbox" name="item_id" invoice="<%=(rsGetRecords.Fields.Item("ID").Value)%>" id="<%=(rsGetOrderDetails2.Fields.Item("OrderDetailID").Value)%>" value="<%=(rsGetOrderDetails2.Fields.Item("OrderDetailID").Value)%>">
			<% end if %>
				<%=(rsGetOrderDetails2.Fields.Item("qty").Value)%>
				<span class="mx-2">|</span>
				<a class="mx-1" href="../productdetails.asp?ProductID=<%=(rsGetOrderDetails2.Fields.Item("ProductID").Value)%>" target="_blank"><%=(rsGetOrderDetails2.Fields.Item("title").Value)%></a>
				<span class="mr-1"><%=(rsGetOrderDetails2.Fields.Item("Gauge").Value)%></span>
				<span class="mr-1"><%=(rsGetOrderDetails2.Fields.Item("Length").Value)%></span>
				<%=(rsGetOrderDetails2.Fields.Item("ProductDetail1").Value)%> (<%=(rsGetOrderDetails2.Fields.Item("PreOrder_Desc").Value)%>)
				<span class="badge badge-info ml-2"><%=(rsGetOrderDetails2.Fields.Item("notes").Value)%></span>
		</div>
	</div><!-- row -->
          <%
.Movenext()
Loop
End With 
%>
</div><!-- container -->
<%
rsGetOrderDetails2.Close()
Set rsGetOrderDetails2 = Nothing
rsGetOrderDetails2_numRows = 0
%>
<div id="comments_<%=(rsGetRecords.Fields.Item("ID").Value)%>">       <br>
          <%=(rsGetRecords.Fields.Item("our_notes").Value)%>
          </div>
          </td>
    </tr>
  <% 
  rsGetRecords.MoveNext()
Wend
%>
</table>

</div>
</body>
<script type="text/javascript" src="../js/jquery-2.1.1.min.js"></script>
<script>
$(document).ready(function(){		
		// Get value for item detail ID from selected checkbox
		$('.checkbox_item_id').click(function() {
		var idd= $(this).attr('id');
		var explode = idd.split('_');
		var invoice_id = $(this).attr('invoice');
		var explode_invoice_id = invoice_id.split('_');
				   $.ajax({
				   type: "POST",
				   url: "set-anodized-order-status.asp?completed=yes&id=" + explode[0] + "&invoice=" + explode_invoice_id[0] + "",
				   success: function(data)
				   {
						$('#item_block_' + explode[0]).addClass("gray-text");
						$('#' + explode[0]).hide();
						$('.bo_' + explode[0]).hide();
			//		   $('#item_block_' + explode[0]).hide();
			//		   $('#comments_' + explode_invoice_id[0]).hide();
				   }
				 });
		});
				 
});
</script>
</html>
<%
rsGetRecords.Close()
Set rsGetRecords = Nothing
%>
