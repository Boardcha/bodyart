<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRecords.Source = "SELECT *  FROM sent_items  WHERE shipped = 'On order' ORDER BY date_order_placed ASC"
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.LockType = 1 'Read-only records
rsGetRecords.Open()
rsGetRecords_numRows = 0


' Get custom order companies
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT name FROM TBL_Companies WHERE preorder_status = 1"
Set rsGetCompanies = objCmd.Execute()
%>
<%
Dim rsEdit__MMColParam
rsEdit__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsEdit__MMColParam = Request.QueryString("ID")
End If
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Ship out custom orders</title>
<style>
	.row:hover select  {background-color:#6FA59A}
	.row:hover input[type="checkbox"]{outline:2px solid #6FA59A;outline-offset: -2px;}
</style>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">

<h4>Custom orders review to ship out</h4>

<div class="form-inline">
<select class="form-control form-control-sm" name="brand_filter" id="brand_filter">
  <option value="-">Filter by brand</option>
  
  <% while not rsGetCompanies.eof %>
  <option value="<%= replace(rsGetCompanies("name"), " ", "_") %>"><%= rsGetCompanies.Fields.Item("name").Value %></option>
  <% rsGetCompanies.movenext()
  wend
  %>
</select>
</div>

<table class="table table-striped table-borderless table-hover mt-3">
  <% 
While NOT rsGetRecords.EOF
%>
  <tr class="row_items"> 
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
rsGetOrderDetails2.Source = "SELECT OrderDetailID, qty, title, ProductDetail1, PreOrder_Desc, notes, backorder, ProductID, Gauge, Length, brandname, item_received, picture FROM dbo.QRY_OrderDetails WHERE ID = " & rsGetRecords.Fields.Item("ID").Value & " AND customorder = 'yes' AND item_ordered = 1 ORDER BY item_received ASC"
rsGetOrderDetails2.CursorLocation = 3 'adUseClient
rsGetOrderDetails2.LockType = 1 'Read-only records
rsGetOrderDetails2.Open()
%>
<div class="container">
<%
Do While Not.Eof

	received = ""
if rsGetOrderDetails2.Fields.Item("item_received").Value = 1 then
	received = "yes"
end if
%>
	<div class="row h-100 my-2 item_block_brand_<%= replace(rsGetOrderDetails2("brandname"), " ", "_") %>">
		<div class="col-2 form-inline my-1 item_block" id="item_block_<%=(rsGetOrderDetails2.Fields.Item("OrderDetailID").Value)%>">
				<% if received = "" then %>
				 	<% if rsGetOrderDetails2.Fields.Item("backorder").Value = 0 then %>
						  
						  <select name="backorder" class="form-control form-control-sm mr-3 backorder" style="height:auto;padding:0" data-id="<%=(rsGetOrderDetails2.Fields.Item("OrderDetailID").Value)%>">
								<option disabled="disabled" selected="selected">Backorder item...</option>
								<option value="bo-preorder-standard">Standard BO</option>
								<option value="bo-preorder-specs">Spec issue</option>
								<option value="bo-preorder-discontinued">Discontinued</option>
						  </select>
						  <a class="badge badge-warning disableClick bo-show-<%=(rsGetOrderDetails2.Fields.Item("OrderDetailID").Value)%>" style="display:none">Backordered</a>
						  <% else %>
						  <a class="badge badge-warning disableClick">Backordered</a>
						<% end if %>
				<% end if %>         
				
		</div>
		<div class="col my-auto <% if received <> "" then %>small<% end if %>">
			<% if received = "" then %>
				<input class="mr-2 checkbox_item_id" type="checkbox" name="item_id" invoice="<%=(rsGetRecords.Fields.Item("ID").Value)%>" id="<%=(rsGetOrderDetails2.Fields.Item("OrderDetailID").Value)%>" value="<%=(rsGetOrderDetails2.Fields.Item("OrderDetailID").Value)%>">
			<% end if %>
				<%=(rsGetOrderDetails2.Fields.Item("qty").Value)%>
				<img class="pull-left mr-2" style="height:50px;width:50px" src="http://bodyartforms-products.bodyartforms.com/<%= rsGetOrderDetails2("picture") %>">
				<span class="font-weight-bold mr-2"><%=(rsGetOrderDetails2.Fields.Item("brandname").Value)%></span>
				<a class="mx-1" href="../productdetails.asp?ProductID=<%=(rsGetOrderDetails2.Fields.Item("ProductID").Value)%>" target="_blank"><%=Replace((rsGetOrderDetails2.Fields.Item("title").Value), "CUSTOM ORDER", "")%></a>
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
	
        $('#brand_filter').change(function(){
			var brand_value = $('#brand_filter').val();
			$('.row_items').hide();
		//	$('.row_item').not('.item_block_brand_' + brand_value).hide();
			$('.item_block_brand_' + brand_value).closest('tr').show();
			console.log('Brand: ' + brand_value);
        });
		
		
		// Get value for item detail ID from selected checkbox
		$('.checkbox_item_id').click(function() {
		var idd= $(this).attr('id');
		var explode = idd.split('_');
		var invoice_id = $(this).attr('invoice');
		var explode_invoice_id = invoice_id.split('_');
				   $.ajax({
				   type: "POST",
				   url: "preorders/status_set-received.asp?received=yes&id=" + explode[0] + "&invoice=" + explode_invoice_id[0] + "",
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
		

		// Backorder item
		$('.backorder').change(function() {
			var bo_id = $(this).attr('data-id');
			var bo_type = $(this).val();
			$(this).hide();

			 $.ajax({
			 type: "POST",
			 url: "preorders/status_set-received.asp",
			 data: {backorder: "yes", id: bo_id, type: bo_type}
			 })
			.done(function(msg) {
				$(".bo-show-" + bo_id).show();
			
			});
		});

		 
});
</script>
</html>
<%
rsGetRecords.Close()
Set rsGetRecords = Nothing
%>
