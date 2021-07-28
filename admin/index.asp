<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


' set cookie to show live/sandbox mode message only for admin users
Response.Cookies("adminuser") = "yes"
Response.Cookies("adminuser").Path = "/"
Response.Cookies("adminuser").Expires =  DATE + 300
					
set cmd_rsGetRecords = Server.CreateObject("ADODB.command")
cmd_rsGetRecords.ActiveConnection = DataConn
cmd_rsGetRecords.CommandText = "SELECT Count(*) AS Total_ToReview FROM sent_items WHERE shipped = 'Review'"
Set rsGetRecords = cmd_rsGetRecords.Execute()

set cmd_BOReviews = Server.CreateObject("ADODB.command")
cmd_BOReviews.ActiveConnection = DataConn
cmd_BOReviews.CommandText = "SELECT Count(*) AS Total_BOReviews FROM dbo.QRY_OrderDetails WHERE BackorderReview = 'Y'"
Set rsBOReviews = cmd_BOReviews.Execute()

set Cmd_rsGetTotalOrder = Server.CreateObject("ADODB.command")
Cmd_rsGetTotalOrder.ActiveConnection = DataConn
Cmd_rsGetTotalOrder.CommandText = "SELECT Count(*) AS Total_ToShip FROM sent_items  WHERE ship_code = 'paid' AND (shipped = 'Pending shipment' OR shipped = 'SHIPPING BACKORDER' OR shipped = 'RETURN ENVELOPE' OR shipped = 'RESHIP PACKAGE')"
Set rsGetTotalOrder = Cmd_rsGetTotalOrder.Execute()

Set rsGetPurchaseOrders_cmd = Server.CreateObject ("ADODB.Command")
rsGetPurchaseOrders_cmd.ActiveConnection = DataConn
rsGetPurchaseOrders_cmd.CommandText = "SELECT * FROM dbo.TBL_PurchaseOrders WHERE po_hide = 0 AND Received = 'N' AND CAST(DateOrdered AS date) > CAST(GETDATE()-150 AS date) ORDER BY DateOrdered DESC" 
Set rsGetPurchaseOrders = rsGetPurchaseOrders_cmd.Execute

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM TBL_Barcodes_SortOrder" 
Set rs_getsections = objCmd.Execute()
%>


<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Administration</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
	<div class="container-fluid p-0">
		<div class="row">
		  <div class="col">
			<div class="card">
				<div class="card-header">
				  <h4>Order management</h4>
				</div>
				<div class="card-body">
					<ul>
						<li>
							<a href="paid_orders.asp">Orders to be shipped (FINAL STEP)<span class="ml-2 badge badge-danger"><%=(rsGetTotalOrder.Fields.Item("Total_ToShip").Value)%></span></a>
						</li>
						<li>
							<a href="review_orders.asp">Review orders to ship<span  class="ml-2 badge badge-danger"><%=(rsGetRecords.Fields.Item("Total_ToReview").Value)%></span></a>
						</li>
						<%  
						If var_access_level <> "Packaging"  then  %>
						<li>
							<a href="review_problemorders.asp">Approve backorders<span  class="ml-2 badge badge-danger"><%=(rsBOReviews.Fields.Item("Total_BOReviews").Value)%></span></a>
						</li>
						<% end if  %>	
						</ul>
				</div>
			  </div>
		  </div>
		  <div class="col"></div>
		</div>
	  </div>
	
<div class="container-fluid p-0 mt-4">
	<div class="row">
	  <div class="col">
		<div class="card">
			<div class="card-header">
				<h4>Order search</h4>
			</div>
			<div class="card-body">
				<form  class="form-inline" name="invoice_search" action="invoice.asp" method="post">
					<div class="form-group">
						<label for="search-invoice">Invoice #</label>
						<input  class="form-control ml-2" name="invoice_num" type="text" id="search-invoice" size="10">
					  </div>
					<button class="btn btn-secondary ml-3" type="submit">Search</button>
			</form>
				<form  class="form-inline"  name="srch_email" method="get" action="order history.asp">
					<label for="search-email">Email</label>
					<input  class="form-control ml-2" name="var_email" id="search-email" type="text"  size="30">
					<button class="btn btn-secondary ml-3" type="submit">Search</button>
				</form  class="form-inline">
				<form  class="form-inline" name="website_search" method="get" action="order history.asp">
					<label for="search-first">First</label>
					<input class="form-control ml-2" name="var_first" id="search-first" type="text"  size="10">
					<label class="ml-4" for="search-last">Last</label>
					<input class="form-control ml-2" name="var_last" id="search-last" type="text"  size="10" >
					<button class="btn btn-secondary ml-3" type="submit">Search</button>
				</form  class="form-inline">
			</div>
		  </div>
	  </div>
	  <div class="col">
		<div class="card">
			<div class="card-header">
				<h4>Product search</h4>
			</div>
			<div class="card-body">
				<form class="form-inline" name="product_search" action="product-edit.asp" method="get">
					<label for="search-productid">Product #</label>
						<input class="form-control ml-2" name="ProductID" type="text" id="search-productid" size="10">
						<button class="btn btn-secondary ml-3" type="submit">Search</button>
			  </form>

			  <form class="form-inline" name="invoice_search" action="SearchDetailID.asp" method="post">
					<label for="search-detailid">Detail ID #</label>
						<input class="form-control ml-2" name="DetailID" type="text" id="search-detailid" size="10">
						<button class="btn btn-secondary ml-3" type="submit">Search</button>
			  </form>
			 
			  <form class="form-inline" name="location_search" action="location_search.asp" method="post">
					<label for="search-section">Section</label>
						<select class="form-control ml-2" name="section" id="search-section">
							<% While NOT rs_getsections.EOF %>                          
							<option value="<%=(rs_getsections.Fields.Item("ID_Description").Value)%>"><%=(rs_getsections.Fields.Item("ID_Description").Value)%></option>
						  <% 
						  rs_getsections.MoveNext()
						  Wend
						  %> 
					  </select>
					
					<label class="ml-3" for="search-location">Location #</label>
						<input class="form-control ml-2" name="location" type="text"  id="search-location" size="10">
						<button class="btn btn-secondary ml-3" type="submit">Search</button>
			  </form>
			</div>
		  </div>
	  </div>
	</div>
  </div>



	

<%  
	If var_access_level <> "Packaging"  then  %>
	<h4 class="mt-4">
		Current orders summary
	</h4>
	<table class="table table-striped">
		<tbody>
	<% 
While NOT rsGetPurchaseOrders.EOF 
%>
<tr>

	<td>
 
    	<%= FormatDateTime(rsGetPurchaseOrders.Fields.Item("DateOrdered").Value,2)%> <a class="ml-4" href="inventory/view_order.asp?ID=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>"><%=(rsGetPurchaseOrders.Fields.Item("Brand").Value)%></a>
	</td>
</tr>
              <% 

  rsGetPurchaseOrders.MoveNext()
Wend
%>
</tbody>
</table>
<% end if %>
</div>
</body>
</html>
<%
rsGetUser.Close()
Set rsGetUser = Nothing
%>