<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
If var_access_level = "Manager" or var_access_level = "Admin" or var_access_level = "Customer service" or user_name = "Melissa" then

set objcmd = Server.CreateObject("ADODB.command")
objcmd.ActiveConnection = DataConn
objcmd.CommandText = "SELECT Count(*) AS Total_Backorders FROM QRY_Backorders2 WHERE shipped <> N'SHIPPING BACKORDER'  AND customorder <> N'yes' AND (QtyInStock > qty)"
Set rsGetBackorders = objcmd.Execute()

set objcmd = Server.CreateObject("ADODB.command")
objcmd.ActiveConnection = DataConn
objcmd.CommandText = "SELECT Count(*) AS Total_ProblemOrders FROM QRY_ErrorsOnReview"
Set rsGetProblemOrders = objcmd.Execute()

set objcmd = Server.CreateObject("ADODB.command")
objcmd.ActiveConnection = DataConn
objcmd.CommandText = "SELECT Count(*) AS Total_Over150 FROM vw_sum_order_orderhistory WHERE over_150 = 1 AND shipped <>  'CUSTOM ORDER APPROVED' AND shipped <> 'On Order' AND shipped <> 'ON HOLD'"
Set rsOrderOver150 = objcmd.Execute()
%>
<title>Customer Service | Admin</title>

<html>
    <body>
    <!--#include virtual="/admin/admin_header.asp"-->
<style>
    a {color:black}
</style>
    <div class="p-2">
    <h4 class="mb-3">Customer Service Dashboard</h4>

<div class="container-fluid p-0">
    <div class="row no-gutters">
      <div class="col-3 pr-2 pb-3">
        <div class="card h-100">
            <div class="card-header">
              <h5>Issues to Handle</h5>
            </div>
            <div class="card-body">
              <a href="/admin/review-orders-over150.asp"><i class="fa fa-angle-right mr-2"></i><% if rsOrderOver150.Fields.Item("Total_Over150").Value > 0 then%><span class="badge badge-danger mr-2"><%= rsOrderOver150.Fields.Item("Total_Over150").Value %></span><% end if %>Review orders over $150</a>
              <br/>
              <a href="/admin/review_problemorders.asp"><i class="fa fa-angle-right mr-2"></i><% if rsGetProblemOrders.Fields.Item("Total_ProblemOrders").Value > 0 then%><span class="badge badge-danger mr-2"><%= rsGetProblemOrders.Fields.Item("Total_ProblemOrders").Value %></span><% end if %>Review problem orders</a>
              <br/>
              
              <a href="/admin/backorders.asp"><i class="fa fa-angle-right mr-2"></i><% if rsGetBackorders.Fields.Item("Total_Backorders").Value > 0 then%><span class="badge badge-danger mr-2"><%=rsGetBackorders.Fields.Item("Total_Backorders").Value%></span><% end if %>Notify backorders</a>
              <br/>
              <a href="/admin/review_orders.asp"><i class="fa fa-angle-right mr-2"></i><span  class="mr-2 badge badge-danger"><%= rsGetOrdersToReview("Total_ToReview") %></span>Review orders to ship</a><br>
            </div><!-- card body -->
          </div><!-- card -->
    </div><!-- col -->
        <div class="col-3 px-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Invoice Search</h5>
                </div>
                <div class="card-body">
                    <form class="form-inline" action="/admin/invoice.asp" method="post">
						<input class="form-control form-control" type="text" name="invoice_num" placeholder="Inovice #">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>
					<form class="form-inline" action="/admin/order history.asp" method="get">
						<input class="form-control form-control" type="text" name="var_email" placeholder="E-mail">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>	
					<form class="form-inline" name="/admin/website_search" method="get" action="order history.asp">
						<input class="form-control form-control mr-2" name="var_first" type="text" placeholder="First name" size="10">
						<input class="form-control form-control" name="var_last" type="text" placeholder="Last name" size="10" ><br>
						<button class="btn btn-sm btn-secondary mt-1" type="submit">Search</button>
					</form>	
				
                </div><!-- card body -->
              </div><!-- card -->
        </div>
        <div class="col-3 px-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Customer Account Search</h5>
                </div>
                <div class="card-body">
                    <form action="/admin/customer_search.asp" method="get">
						
						<input class="form-control form-control mb-2" name="first" type="text" placeholder="First name">
						
						<input class="form-control form-control mb-2" name="last" type="text" placeholder="Last name">
						
						<input class="form-control form-control mb-2" name="email" type="text" placeholder="E-mail">
							
						<input class="form-control form-control mb-2" name="CustomerID" type="text" placeholder="Customer #">
						<button class="btn btn-sm btn-secondary" type="submit">Search</button>
					</form>
                </div><!-- card body -->
              </div><!-- card -->
        </div>
        <div class="col-3 px-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Other Searches</h5>
                </div>
                <div class="card-body">
					<form class="form-inline" action="/admin/order history.asp" method="get">
						<input class="form-control form-control" name="UPS" type="text" placeholder="UPS tracking #">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>
					<form class="form-inline" action="/admin/invoice.asp" method="post">
						<input class="form-control form-control" name="TransID" type="text" placeholder="Transaction ID #">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>	
                </div><!-- card body -->
              </div><!-- card -->
        </div><!-- col -->
        <div class="col-3 pr-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Gift Certificates</h5>
                </div>
                <div class="card-body">
                    <form class="form-inline mb-2" action="/admin/search_giftcertificate.asp" method="post">
						<input class="form-control form-control" name="GiftCert" type="text" placeholder="Gift certificate code">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>							
                        <a class="mt-2" href="/admin/search_giftcertificate.asp"><i class="fa fa-angle-right mr-2"></i>Search gift certificates</a><br/>
                        <a href="/admin/GiftCerts_Combine.asp"><i class="fa fa-angle-right mr-2"></i>Combine gift certificate</a><br/>
                        <a href="/admin/giftcertificate_add.asp"><i class="fa fa-angle-right mr-2"></i>Add new gift certificate</a>
					</form>
                </div><!-- card body -->
              </div><!-- card -->
        </div><!-- col -->
        <div class="col-3 px-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Misc</h5>
                </div>
                <div class="card-body">
                  <a href="/admin/returns.asp"><i class="fa fa-angle-right mr-2"></i>Orders by status</a>
                  <br/>
                  <a href="/admin/PurchaseOrders.asp"><i class="fa fa-angle-right mr-2"></i>Purchase orders</a>
                  <br/>
                  <a href="/admin/invoice.asp?create-empty-order=yes"><i class="fa fa-angle-right mr-2"></i>Create Empty Invoice</a><br/>
                  <a href="/admin/edit_select.asp?jewelry=save"><i class="fa fa-angle-right mr-2"></i>Saved jewelry</a>
                  <br/>
                  <a href="/admin/one-time-coupons.asp"><i class="fa fa-angle-right mr-2"></i>One time use coupons</a>
                  <br/>
                  <a href="/admin/Gallery_EditPhoto.asp"><i class="fa fa-angle-right mr-2"></i>Move gallery photo</a>
                  <br>
                  <a href="/admin/edits_logs.asp"><i class="fa fa-angle-right mr-2"></i>Edit logs</a><br/>
                </div><!-- card body -->
              </div><!-- card -->
        </div><!-- col -->
    </div><!-- row -->
</div><!-- container -->


        </div><!-- body padding-->
    </body>
</div>



</html>

<%
else
    response.write "Access denied"
end if ' permissions
%>
