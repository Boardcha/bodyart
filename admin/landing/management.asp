<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
If var_access_level = "Manager" or var_access_level = "Admin" then
%>
<title>Management | Admin</title>

<html>
    <body>
    <!--#include virtual="/admin/admin_header.asp"-->
<style>
    a {color:black}
</style>
    <div class="p-2">
    <h4 class="mb-3">Management Dashboard</h4>

        <div class="card mb-3">
            <div class="card-header">
              <h5>REPORTS</h5>
            </div>
            <div class="card-body">
                <div class="container-fluid">
                    <div class="row">
                        <div class="col-2 pb-2">
                            <!--#include virtual="/admin/power-bi/inc-report-sales.inc"-->
                        </div><!-- col -->
                        <div class="col-2 pb-2">
                            <!--#include virtual="/admin/power-bi/inc-potential-revenue.inc"-->
                        </div><!-- col -->
                        <div class="col-2 pb-2">
                            <!--#include virtual="/admin/power-bi/inc-analytics-links.inc"-->
                        </div><!-- col -->
                        <div class="col-2 pb-2">
                            <!--#include virtual="/admin/power-bi/inc-inventory.inc"-->
                        </div><!-- col -->
                        <div class="col-2 pb-2">
                            <!--#include virtual="/admin/power-bi/inc-packing-dept.inc"-->
                        </div><!-- col -->
                        <div class="col-2 pb-2">
                           <!--#include virtual="/admin/power-bi/inc-employee-metrics.inc"-->
                        </div><!-- col -->
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

    
          <div class="container-fluid p-0">
            <div class="row no-gutters">
                <div class="col">
                    <div class="card">
                        <div class="card-header">
                          <h5>Employee management</h5>
                        </div>
                        <div class="card-body">
                            <a href="/admin/edits_logs.asp"><i class="fa fa-angle-right mr-2"></i>Edit logs</a><br/>
                            
                            <a href="/admin/manage_users.asp"><i class="fa fa-angle-right mr-2"></i>Manage admin users</a><br/>
                            <a href="https://docs.google.com/spreadsheets/d/1F4iAbeLonejM3rUlWv_vmfPot3dqs_SshWi3YIb-NvY/edit#gid=0" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">New hire checklist</a><br/>
                            <a href="https://docs.google.com/spreadsheets/d/1jgtGqnbMEoX5-zRf0aatDizJ343JZtnHOJQKpbjx6Ls/edit#gid=0" target="_blank">
                                <img class="mr-1" src="/images/icons/google-drive.png" height="20px">Termination checklist</a>
                        
                        </div><!-- card body -->
                      </div><!-- card -->
                </div>
                <div class="col px-2">
                    <div class="card">
                        <div class="card-header">
                          <h5>Site management</h5>
                        </div>
                        <div class="card-body">
                            <a href="/admin/coupons_manage.asp"><i class="fa fa-angle-right mr-2"></i>Manage coupons</a><br/>
                            <a href="/admin/one-time-coupons.asp"><i class="fa fa-angle-right mr-2"></i>One time use coupons</a><br/>
                            <a href="/admin/sliders/sliders.asp"><i class="fa fa-angle-right mr-2"></i>Manage home page sliders</a><br/>
                            <a href="/admin/manage_shippingmethods.asp"><i class="fa fa-angle-right mr-2"></i>Shipping options</a><br/>
                            <a href="/admin/sandbox.asp"><i class="fa fa-angle-right mr-2"></i>Enable sandbox testing</a><br/>
                            <a href="/admin/duplicate-customer-accounts.asp"><i class="fa fa-angle-right mr-2"></i>Customer duplicate accounts</a><br/>
                                <a href="/admin/shipping-notice.asp"><i class="fa fa-angle-right mr-2"></i>Shipping notice</a><br/>
                                <a href="/admin/countries-manage.asp"><i class="fa fa-angle-right mr-2"></i>Manage countries (front end & back end)</a><br/>
                                <a href="/admin/seo-product-search-manager.asp"><i class="fa fa-angle-right mr-2"></i>SEO Product search manager</a><br/>
                                <a href="/admin/seo-title-description-tagging.asp"><i class="fa fa-angle-right mr-2"></i>SEO Title & description tagging</a>
        <br><br>
        
                                <style>
                                    #autoclave-inner:before, #checkout-cards-inner:before, #checkout-paypal-inner:before {
                                        content: "ON"
                                    }
            
                                    #autoclave-inner:after, #checkout-cards-inner:after, #checkout-paypal-inner:after {
                                        content: "OFF"
                                    }
                                </style>
            
                                        <a href="#">Toggle autoclave option on checkout</a>
                                        <div class="onoffswitch small mb-3">
                                            <input type="checkbox" name="toggle_autoclave" class="onoffswitch-checkbox" id="toggle_autoclave" <%=autoclave_checked%>>
                                            <label class="onoffswitch-label" style="margin-bottom:0;margin-top:3px" for="toggle_autoclave">
                                                <span id="autoclave-inner" class="onoffswitch-inner"></span>
                                                <span class="onoffswitch-switch"></span>
                                            </label>
                                        </div>
                                            
                                        <a href="#">Toggle credit card checkout</a>
                                        <div class="onoffswitch small mb-3">
                                            <input type="checkbox" name="toggle_checkout_cards" class="onoffswitch-checkbox" id="toggle_checkout_cards" <%=checkout_cards_checked%>>
                                            <label class="onoffswitch-label" style="margin-bottom:0;margin-top:3px" for="toggle_checkout_cards">
                                                <span id="checkout-cards-inner" class="onoffswitch-inner"></span>
                                                <span class="onoffswitch-switch"></span>
                                            </label>
                                        </div>		
            
                                        <a href="#">Toggle PayPal checkout</a>
                                        <div class="onoffswitch small mb-3">
                                            <input type="checkbox" name="toggle_checkout_paypal" class="onoffswitch-checkbox" id="toggle_checkout_paypal" <%=checkout_paypal_checked%>>
                                            <label class="onoffswitch-label" style="margin-bottom:0;margin-top:3px" for="toggle_checkout_paypal">
                                                <span id="checkout-paypal-inner" class="onoffswitch-inner"></span>
                                                <span class="onoffswitch-switch"></span>
                                            </label>
                                        </div>	
        
        
                        </div><!-- card body -->
                      </div><!-- card -->
                </div><!-- col -->
                <div class="col-3 pr-2 pb-3">
                    <div class="card">
                        <div class="card-header">
                          <h5>Quick Links</h5>
                        </div>
                        <div class="card-body">
                            <a href="https://docs.google.com/document/d/1-b-Id7gYClynkQPyNYitjaG0mhYQ9ey7FP3b1x6dlrI/" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Manager handbook</a><br/>
                        </div><!-- card body -->
                      </div><!-- card -->
                </div><!-- column-->
                <% If var_access_level = "Admin" then %>
                <div class="col">
                    <div class="card">
                        <div class="card-header">
                          <h5>Ellen & Amanda ONLY</h5>
                        </div>
                        <div class="card-body">
                            <a href="/admin/authnet-batches.asp"><i class="fa fa-angle-right mr-2"></i>Batches</a>                
                        </div><!-- card body -->
                      </div><!-- card -->
                </div><!-- col -->
                <% end if %>
            </div><!-- row -->
        </div><!-- container -->


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
            <iframe class="" id="load-report" width="1900px" height="740px" frameborder="0" allowFullScreen="true" scrolling="no" src=""></iframe>
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

    // Toggles
	$("#toggle_autoclave, #toggle_checkout_cards, #toggle_checkout_paypal").on("click", function () {
		$.ajax({
			method: "POST",
			url: "/admin/toggle.asp",
			data: {toggleItem: $(this).attr("id"), isChecked: $(this).is(":checked")}
		})		
	});	
</script>
<%
else
    response.write "Access denied"
end if ' permissions
%>
