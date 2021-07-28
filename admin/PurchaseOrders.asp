<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


set cmd_rsGetWaitingList = Server.CreateObject("ADODB.command")
cmd_rsGetWaitingList.ActiveConnection = DataConn
cmd_rsGetWaitingList.CommandText = "SELECT Count(*) AS Total_Waiting FROM dbo.QRYTopWaitingListItems WHERE qty >= waiting_qty"
Set rsGetWaitingList = cmd_rsGetWaitingList.Execute()



Dim rsGetPurchaseOrders__MMColParam
rsGetPurchaseOrders__MMColParam = "N"
If (Request("MM_EmptyValue") <> "") Then 
  rsGetPurchaseOrders__MMColParam = Request("MM_EmptyValue")
End If



Dim rsGetPurchaseOrders
Dim rsGetPurchaseOrders_cmd
Dim rsGetPurchaseOrders_numRows

Set rsGetPurchaseOrders_cmd = Server.CreateObject ("ADODB.Command")
rsGetPurchaseOrders_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetPurchaseOrders_cmd.CommandText = "SELECT * FROM dbo.TBL_PurchaseOrders where po_hide = 0 ORDER BY Received ASC, PurchaseOrderID DESC" 
rsGetPurchaseOrders_cmd.Prepared = true
rsGetPurchaseOrders_cmd.Parameters.Append rsGetPurchaseOrders_cmd.CreateParameter("param1", 200, 1, 1, rsGetPurchaseOrders__MMColParam) ' adVarChar

Set rsGetPurchaseOrders = rsGetPurchaseOrders_cmd.Execute
rsGetPurchaseOrders_numRows = 0



Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetPurchaseOrders_numRows = rsGetPurchaseOrders_numRows + Repeat1__numRows

%>

<html>
<head>
<title>Purchase orders</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">


<% If Not rsGetWaitingList.EOF Or Not rsGetWaitingList.BOF Then %>

<div class="card mt-3">
  <div class="card-header h5">
    Total items in stock that can be notified:&nbsp;<span id="total-waiting"><%=(rsGetWaitingList.Fields.Item("Total_Waiting").Value)%></span>
  </div>
  <div class="card-body">
    <a class="btn btn-purple text-light" id="notify-waiting-list">Notify customers</a>&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;<a href="waitinglist_compare.asp" target="_top">View list</a></strong></strong>
  </div>
</div>

   <% End If ' end Not rsGetWaitingList.EOF Or NOT rsGetWaitingList.BOF %>



<div class="card my-3">
  <div class="card-header h5">
    Alter label queries
  </div>
  <div class="card-body">
    <!--#include file="labels/inc-update-label-queries.asp" -->
  </div>
</div> 


      <% If Not rsGetPurchaseOrders.EOF Or Not rsGetPurchaseOrders.BOF Then %>  

    <table class="table table-striped table-borderless table-hover">
<thead class="thead-dark">
	<tr>
              <th class="align-middle">Delete</th>
              <th class="align-middle">Date placed</th>
			        <th class="align-middle">Amount</th>
              <th class="align-middle" >Company</th>
              <th class="align-middle">Received</th>
            </tr>
</thead>
              <% 
While NOT rsGetPurchaseOrders.EOF
%>
                <tr id="<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>">
                  <td class="align-middle">
                  <button type="button" class="btn btn-sm btn-danger delete_po" data-po_id="<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>"><i class="fa fa-trash-alt"></i></button>
                </td>
                  <td class="align-middle"><%= FormatDateTime(rsGetPurchaseOrders.Fields.Item("DateOrdered").Value,2)%>&nbsp;&nbsp;&nbsp;#<a href="barcodes_modifyviews.asp?ID=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>&type=Order"><%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%></a>
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
                  </td>
				  <td class="align-middle">
					<%= FormatCurrency(rsGetPurchaseOrders.Fields.Item("po_total").Value, -1, -2, -0, -2) %>
				  </td>
                  <td class="align-middle"><p><strong><%=(rsGetPurchaseOrders.Fields.Item("Brand").Value)%></strong>
                    <% if (rsGetPurchaseOrders.Fields.Item("Received").Value) = "N" then %><br>

					<a href="inventory/view_order.asp?ID=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>">
					View order</a>&nbsp;&nbsp;|&nbsp;&nbsp;
					<a href="inventory_po_process.asp?po_id=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>&new=yes">
					Process order</a>&nbsp;&nbsp;| &nbsp;&nbsp;
					<a href="inventory/create_csv_po.asp?po_id=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>">
					Download .CSV (Excel)</a>
					&nbsp;&nbsp;|&nbsp;&nbsp;
					
					<a href="barcodes_modifyviews.asp?ID=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>&type=new_po_system">Update barcode query</a>
                    <% end if %>
                  </td>
                  <td class="align-middle"><% if (rsGetPurchaseOrders.Fields.Item("Received").Value) = "Y" then %>
                    <span class="badge badge-success">Yes</span> <%=(rsGetPurchaseOrders.Fields.Item("DateReceived").Value)%>
                    <% else %><span class="badge badge-danger">No</span><% end if %></td>
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

    // Notify customers on waiting list
    $(document).on("click", "#notify-waiting-list", function(event){
      $('#notify-waiting-list').html('<i class="fa fa-spinner fa-2x fa-spin"></i>');

      $.ajax({
      dataType: "json",
      url: "/admin/WaitingList_Notify.asp"
      })
      .done(function(json, msg ) {
          $('#total-waiting').html(json.total);
          $('#notify-waiting-list').html('Notify customers');
      })
      .fail(function(json, msg) {
         alert("Failed to notify customers, code error");
      });
  });
</script>
<%
rsGetPurchaseOrders.Close()
Set rsGetPurchaseOrders = Nothing
%>
