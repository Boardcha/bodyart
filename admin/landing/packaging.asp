<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set Cmd_rsGetTotalOrder = Server.CreateObject("ADODB.command")
Cmd_rsGetTotalOrder.ActiveConnection = DataConn
Cmd_rsGetTotalOrder.CommandText = "SELECT Count(*) AS Total_ToShip FROM sent_items  WHERE ship_code = 'paid' AND (shipped = 'Pending shipment' OR shipped = 'SHIPPING BACKORDER' OR shipped = 'RETURN ENVELOPE' OR shipped = 'RESHIP PACKAGE')"
Set rsGetTotalOrder = Cmd_rsGetTotalOrder.Execute()
%>
<title>Packaging | Admin</title>

<html>
    <body>
    <!--#include virtual="/admin/admin_header.asp"-->
<style>
    a {color:black}
</style>
    <div class="p-2">
    <h4 class="mb-3">Packaging Dashboard</h4>

<div class="container-fluid p-0">
    <div class="row no-gutters">
        <div class="col-3 px-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Morning Batch Printing</h5>
                </div>
                <div class="card-body">
                    <a href="/admin/review_orders.asp"><i class="fa fa-angle-right mr-2"></i><span  class="mr-2 badge badge-danger"><%= rsGetOrdersToReview("Total_ToReview") %></span>Review orders to ship</a><br>
                    <a href="/admin/batch-shipping.asp"><i class="fa fa-angle-right mr-2"></i>Print Shipping Forms</a><br/>

                    <a href="/admin/paid_orders.asp"><i class="fa fa-angle-right mr-2"></i><span class="mr-2 badge badge-danger"><%=(rsGetTotalOrder.Fields.Item("Total_ToShip").Value)%></span>Orders to be shipped (FINAL STEP)</a><br>


                    <a href="/admin/insertUPSTracking_numbers.asp"><i class="fa fa-angle-right mr-2"></i>Import UPS #'s</a><br>
                </div><!-- card body -->
              </div><!-- card -->
        </div><!-- col -->
        <div class="col-3 pr-2 pb-3">
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
        <div class="col-3 pr-2 pb-3">
          <div class="card h-100">
            <div class="card-header">
              <h5>Product Search</h5>
            </div>
            <div class="card-body">
              <form class="form-inline" name="invoice_search" action="/admin/product-edit.asp" method="get">
                <input class="form-control form-control" name="ProductID" type="text" placeholder="Product ID #">
                <button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
              </form>
              <form class="form-inline" name="detailid_search" action="/admin/SearchDetailID.asp" method="post">
                <input class="form-control form-control" name="DetailID" type="text" placeholder="Detail ID #">
                <button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
              </form>
              <form class="form-inline" name="location_search" action="/admin/location_search.asp" method="post">
                <select class="form-control mb-1" name="section" id="search-section">
                  <% While NOT rs_getsections.EOF %>                          
                  <option value="<%=(rs_getsections.Fields.Item("ID_Description").Value)%>"><%=(rs_getsections.Fields.Item("ID_Description").Value)%></option>
                  <% 
                  rs_getsections.MoveNext()
                  Wend
                  %> 
                </select>
                <input class="form-control form-control" name="location" type="text" placeholder="Location #">
                <button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
              </form>
            </div><!-- card body -->
            </div><!-- card -->
        </div><!-- column -->
        <div class="col-3 px-2 pb-3">
          <div class="card h-100">
              <div class="card-header">
                <h5>Manuals</h5>
              </div>
              <div class="card-body">
                <a href="https://docs.google.com/document/d/1Hxgg8Xfi9kH1tuu5vTS_9kZ4z806ADFj2TPKYac8q28/" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Packer handbook</a><br/>
                <a href="https://docs.google.com/document/d/1YDJ1B9LrZ8r6kMcXQY3wgQ0yzSC_iHd68wUuhrE0hDQ/" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Printing orders</a><br/>
                <a href="https://docs.google.com/document/d/1miw4HWpPNfhGSMG2GtI2k_AqiQyHmkAUhPfIrEPgS70/" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Packaging orders</a><br/>
                <a href="https://docs.google.com/document/d/1AmX9Y5OB6Gy_v_EuPTrdNIl3WsmERy-JK7sroNydtEc/" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Print training</a><br/>
                <a href="https://docs.google.com/document/d/1e9mHO6ZIWFKqswdSXbzMZ_RfZHhJfo4y989ZX5kUNDw/" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Return mailers</a><br/>
                <a href="https://docs.google.com/document/d/1eYfo1Q6aJ1PlbMnste5rJFWyqGq3hrQim-fIrRQn9hs/" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Autoclave maintenance</a><br/>
                <a href="https://docs.google.com/document/d/11iQkTsdzyZhuk2NPcZr94ejdptP_YwGkQfV5tJcnCXw/" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Making distilled water</a><br/>
                </form>
              </div><!-- card body -->
            </div><!-- card -->
      </div><!-- col -->
    </div><!-- row -->
</div><!-- container -->

<div class="card">
  <div class="card-header">
    <h5 class="d-inline">Update Labels</h5><a class="ml-4" href="/admin/update-labels.asp"><i class="fa fa-angle-right mr-2"></i>Update labels</a>
  </div>
  <div class="card-body">
    <!--#include virtual="/admin/labels/inc-update-label-queries.asp" -->
  </div><!-- card body -->
</div><!-- card -->

        </div><!-- body padding-->
    </body>
</div>



</html>
