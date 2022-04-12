<%@LANGUAGE="VBSCRIPT" %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if Request.querystring("Company") <> "" then
	var_brand = Request.querystring("Company")
else
	var_brand = "Select vendor"
end if

set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT InvoiceID, ProductID, DetailID, qty, PreOrder_Desc, detail_code, title, ProductDetail1, OrderDetailID, Gauge, Length, picture, customer_first FROM dbo.QRY_OrderDetails WHERE customorder = 'yes' AND (shipped = 'CUSTOM ORDER IN REVIEW' or shipped = 'ON ORDER') AND item_ordered = 0 AND brandname = ? ORDER BY InvoiceID ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("brandname",200,1,100, var_brand ))
Set rsGetPreorders = objCmd.Execute()

' Get custom order companies
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT name FROM TBL_Companies WHERE preorder_status = 1"
Set rsGetCompanies = objCmd.Execute()
%>
<html>
<head>
<title>Create custom order purchase order</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h5>
	Create custom order purchase order
</h5> 

<form class="form-inline mb-2">
		<select class="form-control form-control-sm mr-3" name="brandChange" id="brandChange">
			<option value="#">Select company</option>
		<% if var_brand <> "" then %>
			<option value="<%= var_brand %>" selected><%= var_brand %></option>
		<% end if %>
		<% while not rsGetCompanies.eof %>
		<option value="preorder-create-po.asp?Company=<%= rsGetCompanies.Fields.Item("name").Value %>"><%= rsGetCompanies.Fields.Item("name").Value %></option>
		<% rsGetCompanies.movenext()
		wend
		rsGetCompanies.movefirst()
		%>
		</select> 
</form>			



<% If Not rsGetPreorders.EOF Or Not rsGetPreorders.BOF Then %>
			  <table class="table table-sm table-hover mt-4" style="border-collapse:collapse">
				  <thead class="thead-dark">
				<tr>
				  <th width="10%">Invoice</th>
				  <th class="text-center" width="10%">Qty</th>
				  <th width="10%">Code</th>
				  <th width="70%">Description</th>
				</tr>
			</thead>
				<% 
	While NOT rsGetPreorders.EOF
	%>
				  <tr>
					<td>
						<a href="invoice.asp?ID=<%=(rsGetPreorders.Fields.Item("InvoiceID").Value)%>" target="_blank"> <%=(rsGetPreorders.Fields.Item("InvoiceID").Value)%></a>
						<a class="d-block mt-2" href="email_template_send.asp?ID=<%=(rsGetPreorders.Fields.Item("InvoiceID").Value)%>&type=generic">Email <%=(rsGetPreorders.Fields.Item("customer_first").Value)%></a>
					</td>
					<td class="text-center">
						<%=(rsGetPreorders.Fields.Item("qty").Value)%>
					</td>
					<td>
						<%=(rsGetPreorders.Fields.Item("detail_code").Value)%>
					</td>
					<td class="ajax-update">
						<img class="pull-left mr-2" style="height:50px;width:50px" src="http://bodyartforms-products.bodyartforms.com/<%= rsGetPreorders("picture") %>">
						<%=Replace((rsGetPreorders.Fields.Item("title").Value), "CUSTOM ORDER ", "")%>&nbsp;<%=(rsGetPreorders.Fields.Item("gauge").Value)%>&nbsp;<%=(rsGetPreorders.Fields.Item("Length").Value)%>&nbsp;<%=(rsGetPreorders.Fields.Item("ProductDetail1").Value)%><br>
					<div class="form-inline">
					  Update specs: <input class="form-control form-control-sm ml-2 w-75" type="text" name="desc_<%= rsGetPreorders("OrderDetailID") %>" value="<%=(rsGetPreorders.Fields.Item("PreOrder_Desc").Value)%>" data-id="<%= rsGetPreorders("OrderDetailID") %>" data-column="PreOrder_Desc" data-friendly="Custom order specifications" data-int_string="string">
					</div>
					</td>
				  </tr>
				  <% 
	  rsGetPreorders.MoveNext()
	Wend
	%>
	</table>
	<div class="text-center">
		<button class="btn btn-purple" id="create-order">CREATE PURCHASE ORDER</button>
		<span class="text-center ml-2" id="message"></span>
	</div>

<% End If ' end Not rsGetPreorders.EOF Or NOT rsGetPreorders.BOF %>

<% If rsGetPreorders.EOF And rsGetPreorders.BOF Then %>
	<div class="alert alert-danger">No custom orders to review </div>
<% End If ' end rsGetPreorders.EOF And rsGetPreorders.BOF %>

</div>
</body>
</html>
<script type="text/javascript">
	//url to to do auto updating
	var auto_url = "/admin/preorders/ajax-update-custom-orders.asp"

	// URL change menu
	$(document).on("change", '#brandChange', function() { 
		location.href = $(this).val();
	});

	// Create custom order purchase order
	$(document).on("click", "#create-order", function(event){
		var var_brandname = $('#brandChange').val();
		$('#message').html('<i class="fa fa-spinner fa-spin fa-2x"></i>')
		
		$.ajax({
		method: "POST",
		dataType: "json",
		url: "/admin/preorders/ajax-create-custom-order-po.asp",
		data: {brandname: var_brandname}
		})
		.done(function(json,msg ) {
			$('#message').html('<div class="alert alert-success mt-2">PURCHASE ORDER HAS BEEN CREATED<br><a class="btn btn-sm btn-success" href="/admin/inventory/view_order.asp?ID=' + json.purchase_order_id + '" target="_blank">Click here to view order</a></div>')
		})
		.fail(function(json,msg) {
			$('#message').html('<span class="alert alert-danger">ORDER CREATION FAILED</span>')
		});
	});
  </script>
<script type="text/javascript" src="/admin/scripts/generic_auto_update_fields.js"></script>
<script type="text/javascript">
	auto_update(); // run function to update fields when tabbing out of them
</script>
<%
rsGetPreorders.Close()
Set rsGetPreorders = Nothing
Set objCmd = Nothing
%>
