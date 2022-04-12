<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn

objCmd.CommandText = "SELECT TOP (100) PERCENT dbo.sent_items.ID, dbo.sent_items.shipped, dbo.sent_items.customer_first, dbo.sent_items.customer_last, dbo.sent_items.email, dbo.sent_items.country, dbo.sent_items.PackagedBy, dbo.TBL_OrderSummary.ErrorReportDate, dbo.TBL_OrderSummary.ErrorDescription, dbo.TBL_OrderSummary.ErrorOnReview, pulled_by, inventory_issue_description, dbo.sent_items.ship_code, dbo.TBL_OrderSummary.qty, dbo.TBL_OrderSummary.item_price, dbo.TBL_OrderSummary.notes, dbo.ProductDetails.ProductDetail1, dbo.ProductDetails.location, dbo.ProductDetails.Gauge, dbo.ProductDetails.Length, dbo.jewelry.title, dbo.ProductDetails.ProductDetailID, dbo.ProductDetails.BinNumber_Detail, dbo.ProductDetails.wlsl_price, dbo.TBL_OrderSummary.OrderDetailID, dbo.TBL_OrderSummary.ProductID, dbo.TBL_OrderSummary.item_problem, dbo.TBL_OrderSummary.ErrorQtyMissing, dbo.TBL_Barcodes_SortOrder.ID_Description FROM dbo.sent_items INNER JOIN dbo.TBL_OrderSummary ON dbo.sent_items.ID = dbo.TBL_OrderSummary.InvoiceID INNER JOIN dbo.ProductDetails ON dbo.TBL_OrderSummary.DetailID = dbo.ProductDetails.ProductDetailID INNER JOIN dbo.jewelry ON dbo.TBL_OrderSummary.ProductID = dbo.jewelry.ProductID INNER JOIN dbo.TBL_Barcodes_SortOrder ON dbo.ProductDetails.DetailCode = dbo.TBL_Barcodes_SortOrder.ID_Number WHERE (dbo.TBL_OrderSummary.inventory_issue_toggle = 1) ORDER BY dbo.sent_items.ID"
set rsGetRecords = objCmd.Execute()

set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.Open objCmd
%>

<html>
<head>
<title>Review reported inventory issues</title>
<script type="text/javascript" src="../js/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="../js/bootstrap-v4.min.js"></script>
</head>
<body>

<!--#include file="admin_header.asp"-->
<div class="mx-2">
<% If Session("SubAccess") <> "N" then ' DISPLAY ONLY TO PEOPLE WHO HAVE ACCESS TO THIS PAGE %>

<% If NOT rsGetRecords.EOF Then %>
<h5 class="mt-3 mb-2"><%= rsGetRecords.RecordCount %> reported issues</h5>

  <button class="btn btn-sm btn-secondary d-inline-block mb-3" id="update_query_labels" title="Update barcode query for label printing"><i class="fa fa-label fa-lg mr-1"></i> Print requested replacement labels</button>
  <span class="mb-3 ml-1" id="msg-query-update"></span>

<table class="table table-striped table-hover">
	<thead class="thead-dark">
		<tr>
            <th></th>
            <th width="15%">Invoice</th>
			<th width="20%">Item</th>
            <th>Location</th>
            <th>Reported by</th>
			<th width="40%">Reported issue</th>
		</tr>
	</thead>
            <% 
While NOT rsGetRecords.EOF 

if instr(rsGetRecords("inventory_issue_description"), "Print new scanning label") > 0 then
  detailids = detailids & " OR ProductDetailID = " & rsGetRecords("ProductDetailID")
end if
%>
        
	<tr id="row_<%= rsGetRecords.Fields.Item("OrderDetailID").Value %>">
        <td>
            <button class="btn btn-primary btn-sm toggle-off" data-orderdetailid="<%= rsGetRecords.Fields.Item("OrderDetailID").Value %>">Done</button>
        </td>
		<td>
			<strong><a href="invoice.asp?ID=<%= rsGetRecords.Fields.Item("ID").Value %>" class="text-secondary"><%= rsGetRecords.Fields.Item("ID").Value %></strong></a>
		</td>
		<td>
            <a class="text-secondary" href="product-edit.asp?ProductID=<%=(rsGetRecords.Fields.Item("ProductID").Value)%>&info=less"><%=(rsGetRecords.Fields.Item("title").Value)%></a>&nbsp; <%=(rsGetRecords.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetRecords.Fields.Item("Length").Value)%></td>
			<td>
				<%=(rsGetRecords.Fields.Item("ID_Description").Value)%>&nbsp;<%=(rsGetRecords.Fields.Item("location").Value)%>&nbsp;
			<% if (rsGetRecords.Fields.Item("BinNumber_Detail").Value) <> 0 then %>
				(BIN <%=(rsGetRecords.Fields.Item("BinNumber_Detail").Value)%>)
            <% end if %>
			</td>
<td><%=(rsGetRecords.Fields.Item("pulled_by").Value)%></td>
              <td><span class="text-danger"><%=(rsGetRecords.Fields.Item("inventory_issue_description").Value)%></span></td>
              </tr>
            <% 

  rsGetRecords.MoveNext()
  
Wend
%>
          </table>
          <input type="hidden" id="detailids" value="<%= replace(detailids, "OR", "AND",1 , 1) %>">
        <% else ' if there are no records to review %>
		<h5 class="mt-3 mb-2">
			No reported issues
        </h5>
        <% End If ' end rsGetRecords.EOF And rsGetRecords.BOF %>

<% else ' unathorized access error %>
Not accessible
<% end if ' END ACCESS TO PAGE FOR ONLY USERS WHO SHOULD BE ABLE TO SEE IT %>

</div>
<!--#include file="includes/inc_scripts.asp"-->
<script type="text/javascript">
    // Clear error
    $(document).on("click", ".toggle-off", function(){
        var orderdetailid = $(this).attr('data-orderdetailid');

        $.ajax({
            method: "post",
            url: "inventory/toggle-off-inventory-issue.asp",
            data: {orderdetailid: orderdetailid}
            })
            .done(function(msg) {
                $('#row_' + orderdetailid).hide();
            })
            .fail(function(msg) {
                
            })
    })

    	// BEGIN Alter barcode query for item labels
      $(document).on("click", '#update_query_labels', function() { 
      
      $.ajax({
        method: "post",
        url: "/admin/barcodes_modifyviews.asp?type=labels_by_detailid",
        data: {detailids: $('#detailids').val()}
      })
      .done(function() {
        $('#msg-query-update').html('<span class="alert alert-success px-2 py-0"><i class="fa fa-check"></i></span>').show().delay(2500).fadeOut("slow");
      });

    });	// END Alter barcode query for item labels

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
