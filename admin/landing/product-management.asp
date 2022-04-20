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
                    <!--#include virtual="/admin/landing/inc-product-searches.inc"-->
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
                    <a href="/admin/inventory-issues.asp"><i class="fa fa-angle-right mr-2"></i><span class=" badge badge-<% if rsInventoryIssues("total_inventory_issues") > 0 then %>danger<% else %>secondary<% end if %> mr-2"><%= rsInventoryIssues("total_inventory_issues") %></span>Reported inventory issues</a><br/>

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
        <div class="col-3 pr-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Custom Orders</h5>
                </div>
                <div class="card-body">
                    <a href="/admin/custom_orders.asp"><i class="fa fa-angle-right mr-2"></i>Ship out custom orders</a><br/>
                    <a href="/admin/preorder-create-po.asp"><i class="fa fa-angle-right mr-2"></i>Create custom order purchase order</a><br/>
                    <a href="/admin/preorder_emails.asp"><i class="fa fa-angle-right mr-2"></i>E-mails for delays</a><br/>
                    <a href="/admin/one-time-coupons.asp"><i class="fa fa-angle-right mr-2"></i>One time use coupons</a>
                </div><!-- card body -->
              </div><!-- card -->
        </div>
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
        <div class="col-3 px-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Misc</h5>
                </div>
                <div class="card-body">
                    <% 
                    set cmd_rsGetWaitingList = Server.CreateObject("ADODB.command")
                    cmd_rsGetWaitingList.ActiveConnection = DataConn
                    cmd_rsGetWaitingList.CommandText = "SELECT Count(*) AS Total_Waiting FROM dbo.QRYTopWaitingListItems WHERE qty >= waiting_qty"
                    Set rsGetWaitingList = cmd_rsGetWaitingList.Execute()
                    
                    If Not rsGetWaitingList.EOF Then
                    %>                    
                    <div class="mb-4">
                        <a class="btn btn-sm btn-<% if rsGetWaitingList("Total_Waiting") > 0 then %>danger<% else %>secondary<% end if %> mr-2" href="/admin/waitinglist_compare.asp"><span id="total-waiting"><%= rsGetWaitingList("Total_Waiting") %></span> waiting (view)</a><a class="btn btn-sm btn-purple text-light" id="notify-waiting-list">Notify customers</a>
                    </div>
                    <% 
                    End If %>

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
        <div class="col-3 pl-2 pb-3">
          <div class="card h-100">
              <div class="card-header">
                <h5>SOPs & Handbooks</h5>
              </div>
              <div class="card-body">
                <a href="https://docs.google.com/document/d/1wq4u555TkT1SqR7q0gksNTbujsckIP-s1jCnyzylKkE/" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Listing Products</a><br/>
              </div><!-- card body -->
            </div><!-- card -->
      </div><!-- col -->
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
                            <!--#include virtual="/admin/power-bi/inc-report-sales.inc"-->
                        </div><!-- col -->
                        <div class="col-2 pb-2">
                            <!--#include virtual="/admin/power-bi/inc-potential-revenue.inc"-->
                        </div><!-- col -->
                        <% end if ' user_name <> Melissa %>
                        <div class="col-2 pb-2">
                            <!--#include virtual="/admin/power-bi/inc-analytics-links.inc"-->
                        </div><!-- col -->
                        <div class="col-2 pb-2">
                           <!--#include virtual="/admin/power-bi/inc-inventory.inc"-->
                        </div><!-- col -->
                        <% if user_name <> "Melissa" then %>
                        <div class="col-2 pb-2">
                            <!--#include virtual="/admin/power-bi/inc-packing-dept.inc"-->
                        </div><!-- col -->
                        <% end if %>
                        <div class="col-2 pb-2">
                            <!--#include virtual="/admin/power-bi/inc-custom-orders.inc"-->
                        </div><!-- col -->
                        <div class="col-2 pb-2">
                            <!--#include virtual="/admin/power-bi/inc-customers.inc"-->
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
else
    response.write "Access denied"
end if ' permissions
%>
