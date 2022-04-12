<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>

<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

If var_access_level = "Manager" or var_access_level = "Admin" or var_access_level = "Inventory" or var_access_level = "Customer service" then

Dim rsGetPurchaseOrders__MMColParam
rsGetPurchaseOrders__MMColParam = "N"
If (Request("MM_EmptyValue") <> "") Then 
  rsGetPurchaseOrders__MMColParam = Request("MM_EmptyValue")
End If

Set rsGetPurchaseOrders_cmd = Server.CreateObject ("ADODB.Command")
rsGetPurchaseOrders_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetPurchaseOrders_cmd.CommandText = "SELECT *, (SELECT Count(*) AS total_items FROM tbl_po_details WHERE po_orderid = Orders.PurchaseOrderID) AS total_ordered, (SELECT Count(*) AS total_items FROM tbl_po_details WHERE po_received = 1 AND  po_orderid = Orders.PurchaseOrderID) AS total_restocked FROM TBL_PurchaseOrders AS Orders where po_hide = 0 AND po_internal_bulk_pull = 0 ORDER BY Received ASC, PurchaseOrderID DESC" 
rsGetPurchaseOrders_cmd.Prepared = true
rsGetPurchaseOrders_cmd.Parameters.Append rsGetPurchaseOrders_cmd.CreateParameter("param1", 200, 1, 1, rsGetPurchaseOrders__MMColParam) ' adVarChar

Set rsGetPurchaseOrders = rsGetPurchaseOrders_cmd.Execute
%>

<html>
<head>
<title>Purchase orders</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h4>Purchase Orders</h4>

      <% If Not rsGetPurchaseOrders.EOF Or Not rsGetPurchaseOrders.BOF Then %>  

    <table class="table table-striped table-borderless table-hover">
<thead class="thead-dark">
	<tr>
              <% if var_access_level <> "Customer service" then %>          
                <th class="align-middle">Delete</th>
              <% end if %>
              <th class="align-middle">Date placed</th>
              <% if var_access_level <> "Customer service" then %>
                <th class="align-middle">Amount</th>
              <% end if %>
			        <th class="align-middle">Stats</th>
              <th class="align-middle" >Company
                <span class="small"><i class="fa fa-diamond ml-4"></i> = Custom Order</span>
              </th>
              <% if var_access_level <> "Customer service" then %>
                <th class="align-middle">Received</th>
              <% end if %>
            </tr>
</thead>
              <% 
While NOT rsGetPurchaseOrders.EOF

var_percentage_restocked = Round(rsGetPurchaseOrders("total_restocked") / rsGetPurchaseOrders("total_ordered") * 100)
%>
                <tr id="<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>">
                  <% if var_access_level <> "Customer service" then %>
                  <td class="align-middle">
                    
                  <button type="button" class="btn btn-sm btn-info delete_po" data-po_id="<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>"><i class="fa fa-eye-slash"></i></button>
                  
                </td>
                <% end if %>
                  <td class="align-middle"><%= FormatDateTime(rsGetPurchaseOrders.Fields.Item("DateOrdered").Value,2)%>
                    <% if var_access_level <> "Customer service" then %>
                    &nbsp;&nbsp;&nbsp;#<a href="barcodes_modifyviews.asp?ID=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>&type=Order"><%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%></a>
                  <%
                  display_hour = ""
                  display_minutes = ""
                  time_started = ""
                  time_ended = ""
                  get_decimal = ""
                  if rsGetPurchaseOrders.Fields.Item("po_time_started").Value <> "" then 
                    time_started = rsGetPurchaseOrders.Fields.Item("po_time_started").Value
                    time_ended = rsGetPurchaseOrders.Fields.Item("DateOrdered").Value
                    time_difference = DateDiff("n",time_started,time_ended)
                       
                    if time_difference / 60 > 1 then
                      ' ==== IF TIME DIFF HAS A DECIMAL EXPORT TO USER FRIENDLY FORMAT
                      if instr(time_difference / 60, ".") then
                        get_decimal = Split(formatnumber(time_difference / 60, 2), ".")
                        display_hour = get_decimal(0) & " hr"
                        display_minutes = formatnumber((.01 * get_decimal(1)) * 60, 0) & " min"
                      end if
                    else
                      display_minutes = time_difference & " min"
                    end if

                  %>
                    <div>
                        <span class="mr-2"><%= display_hour %></span><%= display_minutes %>
                    </div>
                  <% end if %>
                  <% end if %>
                  </td>
                  <% if var_access_level <> "Customer service" then %>
                  <td class="align-middle">
            
					<%= FormatCurrency(rsGetPurchaseOrders.Fields.Item("po_total").Value, -1, -2, -0, -2) %>
          
				  </td>
          <% end if %>
          <td class="align-middle">
            <%= rsGetPurchaseOrders("total_ordered") %> items ordered<br>
            <%= rsGetPurchaseOrders("total_restocked") %> items restocked <strong><%= var_percentage_restocked %>%</strong><br>
            <% if var_access_level <> "Customer service" then %>
            <a href="/admin/inventory_po_verify.asp?po_id=<%= rsGetPurchaseOrders("PurchaseOrderID") %>">Verify restocks</a>
            <% end if %>
            </td>
            <td class="align-middle">
              <% if rsGetPurchaseOrders("po_type") = "Custom Orders" then %>
                <i class="fa fa-diamond fa-lg mr-1"></i>
              <% end if %>
              <strong><%=(rsGetPurchaseOrders.Fields.Item("Brand").Value)%></strong>
                    <% if (rsGetPurchaseOrders.Fields.Item("Received").Value) = "N" then %><br>

					<a href="inventory/view_order.asp?ID=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>">
					View order</a>
          <% if var_access_level <> "Customer service" then %>
          &nbsp;&nbsp;|&nbsp;&nbsp;
					<a href="inventory_po_process.asp?po_id=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>&new=yes">
					Process order</a>&nbsp;&nbsp;| &nbsp;&nbsp;
					<a href="inventory/create_csv_po.asp?po_id=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>">
					Download Order</a>
					&nbsp;&nbsp;|&nbsp;&nbsp;
					
					<a href="barcodes_modifyviews.asp?ID=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>&type=new_po_system">Update barcode query</a>
                    <% end if %>
                    <% end if %>
                  </td>
                  <% if var_access_level <> "Customer service" then %>
                  <td class="align-middle">
                    
                    <% if (rsGetPurchaseOrders.Fields.Item("Received").Value) = "Y" then %>
                    <span class="badge badge-success">Yes</span> <%=(rsGetPurchaseOrders.Fields.Item("DateReceived").Value)%>
                    <% else %><span class="badge badge-danger">No</span>
                    <span class="badge badge-info pointer po_received" data-po_id="<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>">Set to received</span>
                    <% end if %>
                  
                  </td>
                  <% end if %>
                </tr>
                <% 
  rsGetPurchaseOrders.MoveNext()
Wend
%>
          </table>
      
<% End If ' end Not rsGetPurchaseOrders.EOF Or NOT rsGetPurchaseOrders.BOF %>
      <% If rsGetPurchaseOrders.EOF And rsGetPurchaseOrders.BOF Then %>
        <div class="alert alert-danger">No orders are currently in progress </div>
        <% End If ' end rsGetPurchaseOrders.EOF And rsGetPurchaseOrders.BOF %>
</div>
</body>
</html>
<script type="text/javascript">
  // Delete purchase order
  $(document).on("click", ".delete_po", function(event){
      var po_id = $(this).attr("data-po_id");

      $.ajax({
      method: "POST",
      url: "/admin/inventory/ajax-delete-purchase-order.asp",
      data: {po_id: po_id}
      })
      .done(function(msg ) {
          $('#' + po_id).addClass('table-danger');
          $('#' + po_id).fadeOut('slow');
      })
      .fail(function(msg) {
          alert('FAILED');
      });
  });

  // Update PO to received status 
  $(document).on("click", ".po_received", function(event){
      var po_id = $(this).attr("data-po_id");

      $.ajax({
      method: "POST",
      url: "/admin/inventory/ajax-received-purchase-order.asp",
      data: {po_id: po_id}
      })
      .done(function(msg ) {
          $('#' + po_id).addClass('table-success');
          $('#' + po_id).fadeOut('slow');
      })
      .fail(function(msg) {
          alert('FAILED');
      });
  });
</script>
<%
rsGetPurchaseOrders.Close()
Set rsGetPurchaseOrders = Nothing
%>
<%
else
    response.write "Access denied"
end if ' permissions
%>
