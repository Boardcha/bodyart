<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
If var_access_level = "Manager" or var_access_level = "Admin" or var_access_level = "Inventory" or var_access_level = "Photography" then

set objcmd = Server.CreateObject("ADODB.command")
objcmd.ActiveConnection = DataConn
objcmd.CommandText = "SELECT Count(*) AS total_products FROM jewelry WHERE to_be_pulled = 1 AND pull_completed = 1"
Set rsDiscontinued_Pulled = objcmd.Execute()

set objcmd = Server.CreateObject("ADODB.command")
objcmd.ActiveConnection = DataConn
objcmd.CommandText = "SELECT Count(*) AS total_products FROM jewelry WHERE to_be_pulled = 1 AND pull_completed = 0"
Set rsDiscontinued_ToBePulled = objcmd.Execute()

set objcmd = Server.CreateObject("ADODB.command")
objcmd.ActiveConnection = DataConn
objcmd.CommandText = "SELECT Count(*) AS total_inventory_issues FROM TBL_OrderSummary WHERE inventory_issue_toggle = 1"
Set rsInventoryIssues = objcmd.Execute()
%>
<title>Product Management | Admin</title>

<html>
    <body>
    <!--#include virtual="/admin/admin_header.asp"-->
<style>
    a {color:black}
</style>
    <div class="p-2">
    <h4 class="mb-3">Product Management Dashboard</h4>

<div class="container-fluid p-0">
    <div class="row no-gutters">
        <div class="col-3 pr-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Searches</h5>
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
					<form class="form-inline" name="sku_search" action="/admin/SearchDetailID.asp" method="post">
						<input class="form-control form-control" name="sku" type="text" placeholder="SKU #">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>
                </div><!-- card body -->
              </div><!-- card -->
        </div>
        <div class="col-3 px-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Vendors</h5>
                </div>
                <div class="card-body">
                    <% if user_name <> "Adrienne" and var_access_level <> "Photography" then %>
                    <a href="/admin/inventory-management.asp"><i class="fa fa-angle-right mr-2"></i>Vendor dashboard</a><br/>
                    <a href="/admin/PurchaseOrders.asp"><i class="fa fa-angle-right mr-2"></i>Purchase orders</a><br/>
                    <% end if %>
                    <a href="/admin/add_company.asp"><i class="fa fa-angle-right mr-2"></i>Vendor list</a>
                </div><!-- card body -->
              </div><!-- card -->
        </div><!-- col -->
        <div class="col-3 px-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Products</h5>
                </div>
                <div class="card-body">
                    <a href="/admin/product-edit.asp?add-new-product=yes"><i class="fa fa-angle-right mr-2"></i>Add new product</a><br/>
                    <a href="/admin/edit_select.asp?jewelry=save"><i class="fa fa-angle-right mr-2"></i>Saved jewelry</a><br/>
                    <a href="/admin/available-empty-bins.asp"><i class="fa fa-angle-right mr-2"></i>Available empty bins</a><br/>
                    <a href="/admin/edits_logs.asp"><i class="fa fa-angle-right mr-2"></i>Edit logs</a><br/>

                    <a href="/admin/new-products-sortable.asp"><i class="fa fa-angle-right mr-2"></i>New product sorting</a><br/>
                    <a href="/admin/inventory-issues.asp"><i class="fa fa-angle-right mr-2"></i><span class=" badge badge-danger mr-2"><%= rsInventoryIssues.Fields.Item("total_inventory_issues").Value %></span>Reported inventory issues</a><br/>

                    <a href="/admin/inventory_freeitems.asp"><i class="fa fa-angle-right mr-2"></i>Free item stock</a><br/>
                    <a href="/admin/inventory-count-limited-bin.asp"><i class="fa fa-angle-right mr-2"></i>Limited bin inventory count</a><br/>
                    
                    <a href="/admin/pinned-products.asp"><i class="fa fa-angle-right mr-2"></i>Pinned Products</a><br/>
                    <a href="/admin/Gallery_EditPhoto.asp"><i class="fa fa-angle-right mr-2"></i>Move gallery photo</a>
                </div><!-- card body -->
              </div><!-- card -->
        </div><!-- col -->
        <% if user_name <> "Adrienne" and var_access_level <> "Photography"  then %>
        <div class="col-3 pl-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Discontinued & Sales</h5>
                </div>
                <div class="card-body">
                    <% if rsDiscontinued_Pulled.Fields.Item("total_products").Value > 0 OR rsDiscontinued_ToBePulled.Fields.Item("total_products").Value > 0 then %>
						<a href="/admin/review-pulled-discontinued-items.asp"><i class="fa fa-angle-right mr-2"></i><span class=" badge badge-danger mr-2">To review: <%= rsDiscontinued_Pulled.Fields.Item("total_products").Value %></span><span class=" badge badge-warning mr-2">To be pulled: <%= rsDiscontinued_ToBePulled.Fields.Item("total_products").Value %></span>Review pulled discontinued items</a><br/>
						<% end if %>
                        <% if user_name <> "Jackie" then %>
                    <a href="/admin/inventory-not-moving.asp"><i class="fa fa-angle-right mr-2"></i>Manage inventory not selling</a><br/>
                    <a href="/admin/old-products-sales.asp"><i class="fa fa-angle-right mr-2"></i>Manage sales on old stock</a><br/>
                    <a href="/admin/inventory_clearance.asp"><i class="fa fa-angle-right mr-2"></i>Clearance & Limited</a><br/>
                    <% end if %>
                </div><!-- card body -->
              </div><!-- card -->
        </div><!-- col -->
        <% end if ' permissions on sales %>
        <% if var_access_level <> "Photography" and user_name <> "Adrienne" and user_name <> "Jackie" then %>
        <div class="col-3 pr-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Custom Orders</h5>
                </div>
                <div class="card-body">
                    <a href="/admin/custom_orders.asp"><i class="fa fa-angle-right mr-2"></i>Ship out custom orders</a><br/>
                    <a href="/admin/preorder_approved.asp?Company=Industrial Strength"><i class="fa fa-angle-right mr-2"></i>Approved custom orders</a><br/>
                    <a href="/admin/preorder_review.asp"><i class="fa fa-angle-right mr-2"></i>Review custom orders</a><br/>
                    <a href="/admin/preorder_emails.asp"><i class="fa fa-angle-right mr-2"></i>E-mails for delays</a><br/>
                    <a href="/admin/one-time-coupons.asp"><i class="fa fa-angle-right mr-2"></i>One time use coupons</a>
                </div><!-- card body -->
              </div><!-- card -->
        </div>
        <% end if %>
        <% if var_access_level <> "Photography" and user_name <> "Jackie" then %>
        <div class="col-3 px-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Anodizing</h5>
                </div>
                <div class="card-body">
                    <a href="/admin/anodization-orders.asp"><i class="fa fa-angle-right mr-2"></i><span class=" badge badge-danger mr-2"><%= rsAnodizeCount("total_to_anodize") %></span>Custom orders that need anodizing</a><br/>
					<a href="/admin/inventory-anodize.asp"><i class="fa fa-angle-right mr-2"></i>Anodized products list</a><br/>
					<a href="/admin/available-empty-bins.asp"><i class="fa fa-angle-right mr-2"></i>Available empty bins</a><br/>
					<br>
					<a href="/admin/anodization-management.asp"><i class="fa fa-angle-right mr-2"></i>Colors & voltage pricing</a><br/>
					<br>
                </div><!-- card body -->
              </div><!-- card -->
        </div><!-- col -->
        <% end if %>
        <div class="col-3 px-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Misc</h5>
                </div>
                <div class="card-body">
                    <% if  var_access_level <> "Photography"  then %>
                    <a href="/admin/inventory-bulk-pull-po.asp"><i class="fa fa-angle-right mr-2"></i>Create internal purchase order</a><br/>
                    <% end if %>
                    <% if user_name <> "Adrienne" and user_name <> "Jackie" then %>
                    <a href="/admin/sliders/sliders.asp"><i class="fa fa-angle-right mr-2"></i>Manage home page sliders</a><br/>
                    <% end if %>
                </div><!-- card body -->
              </div><!-- card -->
        </div><!-- col -->
        <% if var_access_level <> "Photography" and user_name <> "Adrienne" and user_name <> "Jackie" then %>
        <div class="col-3 pl-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>3rd Party Selling</h5>
                </div>
                <div class="card-body">
                    <a href="/admin/etsy-manage-inventory.asp?page=1"><i class="fa fa-angle-right mr-2"></i>Etsy stock</a>
                </div><!-- card body -->
              </div><!-- card -->
        </div><!-- col -->
        <% end if %>
    </div><!-- row -->
</div><!-- container -->

<div class="card mb-3">
    <div class="card-header h5">
      Alter label queries
    </div>
    <div class="card-body">
      <!--#include virtual="/admin/labels/inc-update-label-queries.asp" -->
    </div>
  </div> 

<%     If var_access_level = "Admin" or var_access_level = "Manager" or user_name = "Melissa" or user_name = "Nathan" then  %>
        <div class="card mb-3">
            <div class="card-header">
              <h5>REPORTS</h5>
            </div>
            <div class="card-body">
                <div class="container-fluid">
                    <div class="row">
                        <% if user_name <> "Melissa" then %>
                        <div class="col-2 pb-2">
                            <h6 class="border-bottom border-secondary text-center">SALES</h6>
                            <ul class="list-unstyled">
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Orders & sales overview"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSection">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Orders & sales overview
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Shipping carriers"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSectionb0846350843d0cde9228">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Shipping carriers
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Shipping revenue"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSectionba4ce75b01d27f9f9abe">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Shipping revenue
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Payment methods"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSectionc3293043a32b093a128c">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                   Payment methods
                                    </a>
                            </li>
                            <li>
                            <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Wishlist revenue"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSection2eda8b01532747d16079">
                                <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                Wishlist revenue
                                </a>
                                </li>
                                <li>
                                    <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Anodizing revenue"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSection8f49a326e363c433e68c">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        Anodizing revenue
                                        </a>
                                </li>
                                <li>
                                    <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Addons revenue"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSection339116b4a07fec0c33da">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        Addons revenue
                                        </a>
                                </li>
                                <li>
                                    <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Checkout addon sales"  data-report-url="reportEmbed?reportId=b1b71c73-5ea7-41b6-98e7-0f6f18d8bd51&pageName=ReportSection5026a7dbdac7f77784cf">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        Checkout addon sales
                                        </a>
                                </li>
                                <li>
                                    <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Waiting list revenue"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSection94eca4f1a33202e657d8">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        Waiting list revenue
                                        </a>
                                </li>
                                <li>
                                    <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Save for later revenue"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSection68dea6dc51dbc909c598">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        Save for later revenue
                                        </a>
                                </li>
                                <li>
                                    <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Gift certificates"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSection3dbd565056bc0038ac64">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        Gift certificates
                                        </a>
                                </li>
                                <li>
                                    <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Free items"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSection7c15a4140d6f6facd1cf">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        Free items
                                        </a>
                                </li>
                                <li>
                                    <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Return losses"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSectioncc6939f9391d81c1ebf7">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        Return losses
                                        </a>
                                </li>
                                <li>
                                    <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Taxes collected"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSection57a2d0fa08566a3cb9b7">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        Taxes collected
                                        </a>
                                </li>
                                <li>
                                    <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="USA sales"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSectionccdec22c60ecc98573e0">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        USA sales
                                        </a>
                                </li>
                                <li>
                                    <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="International sales"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSectionb60b615e8ed65b97cd82">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        International sales
                                        </a>
                                </li>

                            </ul>
                        </div><!-- col -->
                        <div class="col-2 pb-2">
                            <h6 class="border-bottom border-secondary text-center">POTENTIAL REVENUE</h6>
                            <ul class="list-unstyled">
                                
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Abandoned carts"  data-report-url="reportEmbed?reportId=2fe435c0-68e7-4698-bbf9-9fa6925746e9&pageName=ReportSection8da612ddfee8cd25cd84">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Abandoned carts
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Abandoned carts funnel"  data-report-url="reportEmbed?reportId=2fe435c0-68e7-4698-bbf9-9fa6925746e9&pageName=ReportSectiond4e6bafcae3e7bd89b9c">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Abandoned carts funnel
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Daily abandoned carts"  data-report-url="reportEmbed?reportId=2fe435c0-68e7-4698-bbf9-9fa6925746e9&pageName=ReportSectionebec3316370647bd9e02">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Daily abandoned carts
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Wishlist potential revenue"  data-report-url="reportEmbed?reportId=2fe435c0-68e7-4698-bbf9-9fa6925746e9&pageName=ReportSection94178c9682b7e127a75d">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Wishlist potential revenue
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Save for later potential revenue"  data-report-url="reportEmbed?reportId=2fe435c0-68e7-4698-bbf9-9fa6925746e9&pageName=ReportSectionac38b93470d606e28ed3">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Save for later potential revenue
                                    </a>
                            </li>
                                
                            </ul>
                        </div><!-- col -->
                        <% end if ' user_name <> Melissa %>
                        <div class="col-2 pb-2">
                            <h6 class="border-bottom border-secondary text-center">E-COMMERCE ANALYTICS</h6>
                            <ul class="list-unstyled">
                              
                                <li>
                                    <a href="https://datastudio.google.com/u/0/reporting/b0e4b497-15a9-491d-9134-6eaab615d3a6/page/5auqB" target="_blank">
                                        <img class="mr-1" src="/images/icons/google-data-studio.png" height="20px">
                                        Google dashboard
                                        </a>
                                </li>
                                <li>
                                    <a href="/admin/competitor-research.asp" >
                                        <img class="mr-1" src="/images/icons/google-trends.png" height="20px">
                                        Competitor trends
                                        </a>
                                </li>
                            </ul>
                        </div><!-- col -->
                        <div class="col-2 pb-2">
                            <h6 class="border-bottom border-secondary text-center ">INVENTORY</h6>
                            <ul class="list-unstyled">
                              
                            <li>
                                    <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Incoming vs outgoing"  data-report-url="reportEmbed?reportId=b1b71c73-5ea7-41b6-98e7-0f6f18d8bd51&pageName=ReportSection781f072d6b16235e5969">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        Incoming vs outgoing
                                        </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Wholesale movement"  data-report-url="reportEmbed?reportId=b1b71c73-5ea7-41b6-98e7-0f6f18d8bd51&pageName=ReportSectionc8d3d89f753886d9893b">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Wholesale movement
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Incoming jewelry"  data-report-url="reportEmbed?reportId=b1b71c73-5ea7-41b6-98e7-0f6f18d8bd51&pageName=ReportSection8f6da6f2f8fef955931f">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Incoming jewelry
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Wholesale stock on hand"  data-report-url="reportEmbed?reportId=b1b71c73-5ea7-41b6-98e7-0f6f18d8bd51&pageName=ReportSectionbddd7e9b0e965ec3d95e">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Wholesale stock on hand
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Vendor profits"  data-report-url="reportEmbed?reportId=b1b71c73-5ea7-41b6-98e7-0f6f18d8bd51&pageName=ReportSection40f4fdae8ccf16fc0e2b">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Vendor profits
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Inactive items with qty"  data-report-url="reportEmbed?reportId=b1b71c73-5ea7-41b6-98e7-0f6f18d8bd51&pageName=ReportSectionc36c32481ac846171df1">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Inactive items with qty
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Top selling products"  data-report-url="reportEmbed?reportId=b1b71c73-5ea7-41b6-98e7-0f6f18d8bd51&pageName=ReportSection60e33d96a371a3a00de0">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Top selling products
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Top gauges"  data-report-url="reportEmbed?reportId=b1b71c73-5ea7-41b6-98e7-0f6f18d8bd51&pageName=ReportSection2dabdc6b693ab8a0b22e">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Top gauges
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Backorders"  data-report-url="reportEmbed?reportId=b1b71c73-5ea7-41b6-98e7-0f6f18d8bd51&pageName=ReportSectiondb3e6db04e1c0fb1b26b">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Backorders
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Items rated below 4 stars"  data-report-url="reportEmbed?reportId=b1b71c73-5ea7-41b6-98e7-0f6f18d8bd51&pageName=ReportSection0995471f3d43f1ec43b7">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Items rated below 4 stars
                                    </a>
                            </li>
                            </ul>
                        </div><!-- col -->
                        <% if user_name <> "Melissa" then %>
                        <div class="col-2 pb-2">
                            <h6 class="border-bottom border-secondary text-center">PACKING DEPARTMENT</h6>
                            <ul class="list-unstyled">
                                <li>
                                    <a href="/admin/packing-errors.asp"><i class="fa fa-angle-right mr-2 ml-2"></i>&nbsp;Packer errors</a>
                                </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Orders Overview"  data-report-url="reportEmbed?reportId=15766e9d-3dcc-415f-9d97-f4e9c66fb5e8&pageName=ReportSection" >
                                <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                Orders overview
                                </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Packing Stats"  data-report-url="reportEmbed?reportId=15766e9d-3dcc-415f-9d97-f4e9c66fb5e8&pageName=ReportSectionecb8bec0b16576690ac1" >
                                <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                Packing stats
                                </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Pulling Stats"  data-report-url="reportEmbed?reportId=15766e9d-3dcc-415f-9d97-f4e9c66fb5e8&pageName=ReportSection0cf58b6cd80c063e308a" >
                                <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                Pulling stats
                                </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Items Stats"  data-report-url="reportEmbed?reportId=15766e9d-3dcc-415f-9d97-f4e9c66fb5e8&pageName=ReportSectiona5059775a4468f737145" >
                                <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                Items stats
                                </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Product Restocks"  data-report-url="reportEmbed?reportId=15766e9d-3dcc-415f-9d97-f4e9c66fb5e8&pageName=ReportSectionb09c04b6b8ae803d9e39">
                                <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                Product restocks
                                </a>
                            </li>
                                
                            </ul>
                        </div><!-- col -->
                        <% end if %>
                        <div class="col-2 pb-2">
                            <h6 class="border-bottom border-secondary text-center">CUSTOM ORDERS</h6>
                            <ul class="list-unstyled">
                              
                                <li>
                                    <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Custom orders revenue"  data-report-url="reportEmbed?reportId=c88f768b-44f2-4394-8802-2577a5ef3cdf&pageName=ReportSectionc4da01ec413cf108a641">
                                        <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                        Custom orders revenue
                                        </a>
                                </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Shipping status issues"  data-report-url="reportEmbed?reportId=e1acf509-fb33-4f12-937a-f23c0829fdf1&pageName=ReportSection">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Shipping status issues
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Custom orders with anodizing"  data-report-url="reportEmbed?reportId=e1acf509-fb33-4f12-937a-f23c0829fdf1&pageName=ReportSectionc65c7a14c8c54d39bab7">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    w/ Anodizing
                                    </a>
                            </li>
                               
                            </ul>
                        </div><!-- col -->
                        <div class="col-2 pb-2">
                            <h6 class="border-bottom border-secondary text-center">CUSTOMERS</h6>
                            <ul class="list-unstyled">
                              
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Customer counts"  data-report-url="reportEmbed?reportId=1fbf5d3f-a940-4fbc-83c2-9b8140e3312c&pageName=ReportSectionc9486765338688687030">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Customer counts
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Account creation"  data-report-url="reportEmbed?reportId=1fbf5d3f-a940-4fbc-83c2-9b8140e3312c&pageName=ReportSectionc281d354339d98c1dda8">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Account creation
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Survey replies"  data-report-url="reportEmbed?reportId=1fbf5d3f-a940-4fbc-83c2-9b8140e3312c&pageName=ReportSection">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Survey replies
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Saved searches use"  data-report-url="reportEmbed?reportId=1fbf5d3f-a940-4fbc-83c2-9b8140e3312c&pageName=ReportSectionb93429d7140cc297392e">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Saved searches use
                                    </a>
                            </li>
                            <li>
                                <a href="" class="load-report" data-toggle="modal" data-target="#modal-reports" data-report-title="Registrations & newsletters"  data-report-url="reportEmbed?reportId=2fe435c0-68e7-4698-bbf9-9fa6925746e9&pageName=ReportSection5897cea55cb9f1bdb43f">
                                    <img class="mr-1" src="/images/icons/power-bi.png" height="20px">
                                    Registrations & newsletters
                                    </a>
                            </li>
                            </ul>
                        </div><!-- col -->


                    </div><!-- row -->
                </div><!-- container fluid -->

            
            </div><!-- card body -->
          </div><!-- card -->
<% end if ' reports permissions %>
    



        </div><!-- body padding-->
    </body>
</div>


<!-- Power BI Modal -->
<div class="modal fade" id="modal-reports" tabindex="-1" role="dialog"  aria-labelledby="modal-reports" >
	<div class="modal-dialog modal-dialog-centered mw-100" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title" id="reports-title"></h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<div class="modal-body text-center">
            <iframe class="" id="load-report" width="1340px" height="740px" frameborder="0" allowFullScreen="true" scrolling="no" src=""></iframe>
		</div>
	  </div>
	</div>
</div>
<!-- End Power BI Modal -->

</html>

<script type="text/javascript">
    $(document).on("click", ".load-report", function(event){

		var report_title = $(this).attr("data-report-title");
		var report_url = $(this).attr("data-report-url");

        $('#reports-title').html(report_title);
        $("#load-report").attr("src",'https://app.powerbi.com/' + report_url + '&navContentPaneEnabled=false&filterPaneEnabled=false&autoAuth=true&ctid=06bc9374-9044-4ccb-8d1c-84eb80fc2e89&config=eyJjbHVzdGVyVXJsIjoiaHR0cHM6Ly93YWJpLXVzLWNlbnRyYWwtYS1wcmltYXJ5LXJlZGlyZWN0LmFuYWx5c2lzLndpbmRvd3MubmV0LyJ9');
	
	}); 
</script>
<%
else
    response.write "Access denied"
end if ' permissions
%>
