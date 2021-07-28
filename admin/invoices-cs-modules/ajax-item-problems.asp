<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TOP (100) PERCENT sent_items.ID, sent_items.shipped, sent_items.customer_first, sent_items.customer_last, sent_items.email, sent_items.country, sent_items.PackagedBy, sent_items.ship_code FROM sent_items WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, request.form("invoiceid")))
set rsGetInvoice = objCmd.Execute()

set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TOP (100) PERCENT sent_items.ID, sent_items.shipped, sent_items.customer_first, sent_items.customer_last, sent_items.email, jewelry.picture, sent_items.country, sent_items.PackagedBy, TBL_OrderSummary.ErrorReportDate, TBL_OrderSummary.ErrorDescription, TBL_OrderSummary.ErrorOnReview, sent_items.ship_code, TBL_OrderSummary.qty, TBL_OrderSummary.item_price, TBL_OrderSummary.notes, ProductDetails.ProductDetail1, ProductDetails.location, ProductDetails.Gauge, ProductDetails.qty AS 'amt_instock', ProductDetails.Length, jewelry.title, ProductDetails.ProductDetailID, ProductDetails.BinNumber_Detail, ProductDetails.wlsl_price, TBL_OrderSummary.OrderDetailID, TBL_OrderSummary.ProductID, TBL_OrderSummary.item_problem, TBL_OrderSummary.ErrorQtyMissing, TBL_Barcodes_SortOrder.ID_Description FROM sent_items INNER JOIN TBL_OrderSummary ON sent_items.ID = TBL_OrderSummary.InvoiceID INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE (TBL_OrderSummary.ErrorOnReview = 1) AND ID = ? ORDER BY sent_items.ID"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, request.form("invoiceid")))
set rsGetRecords = objCmd.Execute()

%>

<html>
<body>
	<div class="container mb-3">
		<div class="row">
		  <div class="col">
			<h6 class="d-inline-block mr-5">Invoice #<%= rsGetInvoice.Fields.Item("ID").Value %></h6>
<h6 class="d-inline-block mr-5">Shipping to: <%=(rsGetInvoice.Fields.Item("country").Value)%></h6>
		  </div>
		  <div class="col text-right">
			<a class="btn btn-sm btn-outline-secondary" href="order history.asp?var_first=<%=(rsGetInvoice.Fields.Item("customer_first").Value)%>&var_last=<%=(rsGetInvoice.Fields.Item("customer_last").Value)%>" target="_blank">View history</a>
		  </div>
		</div>
	  </div>

<% If NOT rsGetInvoice.EOF Then %>
<table class="table table-striped  table-bordered table-hover table-sm small">
	<thead class="thead-dark">
		<tr>
			<th width="5%">Problem</th>
			<th class="text-center" width="5%">Qty with issues</th>
			<th class="text-center" width="5%">Qty originally ordered</th>
			<th class="text-center" width="5%">Currently in stock</th>
			<th class="text-center" width="5%">Item location</th>
			<th width="50%">Problem item(s) on order</th>
			<th width="25%">Problem description</th>
		</tr>
	</thead>
            <% 
While NOT rsGetRecords.EOF 
%>
        
	<tr>
		<td class="text-center">
			<% = (rsGetRecords.Fields.Item("item_problem").Value) %>
		</td>
		<td class="text-center ajax-update">
			<input class="form-control form-control-sm" name="qty_<%=(rsGetRecords.Fields.Item("OrderDetailID").Value)%>" value="<%=(rsGetRecords.Fields.Item("ErrorQtyMissing").Value)%>" data-id="<%= rsGetRecords.Fields.Item("OrderDetailID").Value %>" data-detailid="<%= rsGetRecords.Fields.Item("OrderDetailID").Value %>" data-productdetailid="<%= rsGetRecords.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsGetRecords.Fields.Item("ProductID").Value %>" data-column="ErrorQtyMissing" data-friendly="Qty to reship" data-int_string="integer">
		</td>
		<td class="text-center">
			<%=(rsGetRecords.Fields.Item("qty").Value)%>
		</td>
		<td class="text-center">
			<%=(rsGetRecords.Fields.Item("amt_instock").Value)%>
		</td>
		<td>
			<%=(rsGetRecords.Fields.Item("ID_Description").Value)%>&nbsp;<%=(rsGetRecords.Fields.Item("location").Value)%>&nbsp;
			<% if (rsGetRecords.Fields.Item("BinNumber_Detail").Value) <> 0 then %>
				(BIN <%=(rsGetRecords.Fields.Item("BinNumber_Detail").Value)%>)
			<% end if %>
			</td>
			<td>
			<a class="text-secondary" href="product-edit.asp?ProductID=<%=(rsGetRecords.Fields.Item("ProductID").Value)%>&info=less">
				<img src="http://bodyartforms-products.bodyartforms.com/<%=(rsGetRecords.Fields.Item("picture").Value)%>" style="width:50px;height:auto">
				<%=(rsGetRecords.Fields.Item("title").Value)%></a>&nbsp; <%=(rsGetRecords.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetRecords.Fields.Item("Length").Value)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Wholesale: <%= FormatCurrency(rsGetRecords.Fields.Item("wlsl_price").Value * (rsGetRecords.Fields.Item("qty").Value), -1, -2, -2, -2) %>
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
</body>
</html>
<%
DataConn.Close()
%>

