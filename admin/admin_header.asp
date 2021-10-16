<%
set objcmd = Server.CreateObject("ADODB.command")
objcmd.ActiveConnection = DataConn
objcmd.CommandText = "SELECT * FROM TBL_Toggle_Items"
Set rsToggle = objcmd.Execute()
While Not rsToggle.EOF 
	If rsToggle("toggle_item") = "toggle_autoclave" AND rsToggle("value") Then autoclave_checked = "checked"
	If rsToggle("toggle_item") = "toggle_checkout_cards" AND rsToggle("value") Then checkout_cards_checked = "checked"
	If rsToggle("toggle_item") = "toggle_checkout_paypal" AND rsToggle("value") Then checkout_paypal_checked = "checked"
	rsToggle.MoveNext	
Wend
%>

<% ' Set testing user access level
'var_access_level = "Packaging"
%>
	<% if request.cookies("admindarkmode") <> "on" then %> 
		<link href="/CSS/baf.min.css?v=080221" id="lightmode" rel="stylesheet" type="text/css" />
	<% else %>
		<link href="/CSS/baf-dark.min.css?v=080221" id="darkmode" rel="stylesheet" type="text/css" />
	<% end if %>
	<script src="https://use.fortawesome.com/dc98f184.js"></script>

<%
if session("sandbox") = "ON" then
%>
	<div class="bg-warning text-center">
		<h4 class="m-0 p-2">SANDBOX TESTING MODE ON</h4>
	</div>
<% end if ' sandbox
%>

<style>
	.navbar-nav li:hover .dropdown-menu {
    display: block;
}

	.navbar a, .navbar h5 {color: #fff!important}
	.navbar h5 {text-decoration: underline}
</style>

<nav class="navbar navbar-dark bg-dark navbar-expand py-0 px-1" style="z-index:5000">
			<ul class="navbar-nav mr-auto">
				<li class="nav-item">
						<a class="nav-link" href="/admin/index.asp"><i class="fa fa-home mt-1"></i>
						</a>
				</li>
										
<%
'=========== START CUSTOMER SERVICE MENU ========================================
If var_access_level = "Admin" or var_access_level = "Manager" or var_access_level = "Customer service" then

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
objcmd.CommandText = "SELECT Count(*) AS Total_Over150 FROM vw_sum_order_orderhistory WHERE over_150 = 1 AND shipped <>  'Pre-order Approved' AND shipped <> 'On Order' AND shipped <> 'ON HOLD'"
Set rsOrderOver150 = objcmd.Execute()
%>
	<li class="nav-item dropdown position-static border-right border-secondary">
		<a class="nav-link" href="#" id="CSDropdown"
				role="button" data-toggle="dropdown" aria-haspopup="true"
				aria-expanded="false">
				Customer Service</a>

	<div class="dropdown-menu bg-dark w-100 m-0 border-0 rounded-0 pb-5"
				aria-labelledby="CSDropdown">
		<div class="container-fluid px-2">
			<div class="row w-100">
				<div class="col">
					
					<a href="/admin/review-orders-over150.asp">Review orders over $150<% if rsOrderOver150.Fields.Item("Total_Over150").Value > 0 then%> <span class="badge badge-danger ml-2"><%= rsOrderOver150.Fields.Item("Total_Over150").Value %></span><% end if %></a>
					<br/>
					<a href="/admin/review_problemorders.asp">Review problem orders<% if rsGetProblemOrders.Fields.Item("Total_ProblemOrders").Value > 0 then%> <span class="badge badge-danger ml-2"><%= rsGetProblemOrders.Fields.Item("Total_ProblemOrders").Value %></span><% end if %></a>
					<br/>
					
					<a href="/admin/backorders.asp">Notify backorders<% if rsGetBackorders.Fields.Item("Total_Backorders").Value > 0 then%> <span class="badge badge-danger ml-2"><%=rsGetBackorders.Fields.Item("Total_Backorders").Value%></span><% end if %></a>
					<br/>
					<a href="/admin/returns.asp">Orders by status</a>
					<br/>
					<a href="/admin/edit_select.asp?jewelry=save">Saved jewelry</a>
					<br/>
					<a href="/admin/one-time-coupons.asp">One time use coupons</a>
					<br/>
					<a href="/admin/Gallery_EditPhoto.asp">Move gallery photo</a>
					<h5 class="mt-3">Research</h5>
					<a href="/admin/edits_logs.asp">Edit logs</a><br/>
				</div>
				<div class="col">
					<h5>Invoice search</h5>			
					<form class="form-inline" action="invoice.asp" method="post">
						<input class="form-control form-control-sm w-75" type="text" name="invoice_num" placeholder="Inovice #">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>
					<form class="form-inline" action="order history.asp" method="get">
						<input class="form-control form-control-sm w-75" type="text" name="var_email" placeholder="E-mail">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>	
					<form class="form-inline" name="website_search" method="get" action="order history.asp">
						<input class="form-control form-control-sm mr-2" name="var_first" type="text" placeholder="First name" size="10">
						<input class="form-control form-control-sm" name="var_last" type="text" placeholder="Last name" size="10" >
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>		
				</div>
				<div class="col">
					<h5>Customer account search</h5>
					<form action="customer_search.asp" method="get">
						
						<input class="form-control form-control-sm w-75 mb-2" name="first" type="text" placeholder="First name">
						
						<input class="form-control form-control-sm w-75 mb-2" name="last" type="text" placeholder="Last name">
						
						<input class="form-control form-control-sm w-75 mb-2" name="email" type="text" placeholder="E-mail">
							
						<input class="form-control form-control-sm w-75 mb-2" name="CustomerID" type="text" placeholder="Customer #">
						<button class="btn btn-sm btn-secondary" type="submit">Search</button>
					</form>
				</div>
				<div class="col">
					<h5>Other searches:</h5>
					<form class="form-inline" action="order history.asp" method="get">
						<input class="form-control form-control-sm w-75" name="UPS" type="text" placeholder="UPS tracking #">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>
					<form class="form-inline" action="invoice.asp" method="post">
						<input class="form-control form-control-sm w-75" name="TransID" type="text" placeholder="Transaction ID #">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>	

					<h5>Gift certificates</h5>
					<form class="form-inline mb-2" action="search_giftcertificate.asp" method="post">
						<input class="form-control form-control-sm w-75" name="GiftCert" type="text" placeholder="Gift certificate code">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>			
					</form>
					<a href="/admin/search_giftcertificate.asp" class="HomePageLinks">Search gift certificates</a><br/>
					<a href="/admin/GiftCerts_Combine.asp" class="HomePageLinks">Combine gift certificate</a><br/>
					<a href="/admin/giftcertificate_add.asp" class="HomePageLinks">Add new gift certificate</a>
				</div>
			</div>
		</div>
	</div>
</li>
<% end if '=========== END CUSTOMER SERVICE MENU ================================ %>

<%
'=========== START PRODUCT MANAGEMENT MENU ========================================
If var_access_level = "Admin" or var_access_level = "Manager" or var_access_level = "Inventory" then  

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

	<li class="nav-item dropdown position-static border-right border-secondary">
		<a class="nav-link" href="#" id="ProductsDropdown"
				role="button" data-toggle="dropdown" aria-haspopup="true"
				aria-expanded="false">
				Product Management</a>
	<div class="dropdown-menu bg-dark  w-100 m-0 border-0 rounded-0 pb-5"
				aria-labelledby="ProductsDropdown">
		<div class="container-fluid">
			<div class="row w-100">
				<div class="col">
					<h5>Products</h5>
						<a href="/admin/product-edit.asp?add-new-product=yes">Add new product</a><br/>
						<a href="/admin/inventory-issues.asp">Reported inventory issues</a>
							<span class=" badge badge-danger ml-2"><%= rsInventoryIssues.Fields.Item("total_inventory_issues").Value %></span><br/>
						<a href="/admin/inventory-not-moving.asp">Manage inventory not selling</a><br/>
						<a href="/admin/inventory_clearance.asp">Clearance & Limited</a><br/>
						<a href="/admin/inventory_freeitems.asp">Free item stock</a><br/>
						<a href="/admin/inventory-count-limited-bin.asp">Limited bin inventory count</a><br/>
						<% if rsDiscontinued_Pulled.Fields.Item("total_products").Value > 0 OR rsDiscontinued_ToBePulled.Fields.Item("total_products").Value > 0 then %>
						<a href="/admin/review-pulled-discontinued-items.asp">Review pulled discontinued items<span class=" badge badge-danger" style="margin-left:5px">To review: <%= rsDiscontinued_Pulled.Fields.Item("total_products").Value %></span><span class=" badge badge-warning ml-2">To be pulled: <%= rsDiscontinued_ToBePulled.Fields.Item("total_products").Value %></span></a><br/>
						<% end if %>
						<a href="/admin/pinned-products.asp">Pinned Products</a><br/>
						<a href="/admin/Gallery_EditPhoto.asp">Move gallery photo</a>
				</div>
				<div class="col">
					<h5>Vendors</h5>
						<a href="/admin/inventory-management.asp">Vendor dashboard</a><br/>
						<a href="/admin/PurchaseOrders.asp">Purchase orders</a><br/>
						<a href="/admin/add_company.asp">Vendor list</a>

						<h5 class="mt-3">Research</h5>
						<a href="/admin/edits_logs.asp">Edit logs</a><br/>
				</div>
				<div class="col">
					<h5>Searches</h5>
					<form class="form-inline" name="invoice_search" action="product-edit.asp" method="get">
						<input class="form-control form-control-sm w-75" name="ProductID" type="text" placeholder="Product ID #">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>
					<form class="form-inline" name="detailid_search" action="SearchDetailID.asp" method="post">
						<input class="form-control form-control-sm w-75" name="DetailID" type="text" placeholder="Detail ID #">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>
					<form class="form-inline" name="location_search" action="location_search.asp" method="post">
						<input class="form-control form-control-sm w-75" name="location" type="text" placeholder="Location #">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>
					<form class="form-inline" name="sku_search" action="SearchDetailID.asp" method="post">
						<input class="form-control form-control-sm w-75" name="sku" type="text" placeholder="SKU #">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>
					
				</div>
				<div class="col">
					<h5>Etsy</h5>
					<a href="/admin/etsy-manage-inventory.asp?page=1">Etsy stock</a>
				
				</div>
				<div class="col">
					<h5>Anodizing</h5>
					<a href="/admin/inventory-anodize.asp">Anodized products list</a><br/>
					<a href="/admin/available-empty-bins.asp">Available empty bins</a><br/>
					
				</div>
				
			</div>
		</div>
	</div>
</li>
<% end if '=========== END PRODUCT MANAGEMENT MENU ================================ %>


<% '==================== START PRE-ORDERS MENU ====================================

If (user_name = "Melissa" and var_access_level = "Customer service") or var_access_level = "Admin" or var_access_level = "Manager" then  %>
	<li class="nav-item dropdown position-static border-right border-secondary">
		<a class="nav-link" href="#" id="PreOrdersDropdown"
				role="button" data-toggle="dropdown" aria-haspopup="true"
				aria-expanded="false">
				Pre-orders</a>
	<div class="dropdown-menu bg-dark w-100 m-0  border-0 rounded-0 pb-5"
				aria-labelledby="PreOrdersDropdown">
		<div class="container-fluid">
			<div class="row w-100">
				<div class="col">
					<h5>Pre-Orders</h5>
						<a href="/admin/custom_orders.asp">Ship out pre-orders</a><br/>
						<a href="/admin/preorder_approved.asp?Company=Industrial Strength">Approved pre-orders</a><br/>
						<a href="/admin/preorder_review.asp">Review pre-orders</a><br/>
						<a href="/admin/preorder_emails.asp">E-mails for delays</a><br/>
						<a href="/admin/one-time-coupons.asp">One time use coupons</a>
				</div>
				<div class="col">
					<h5>Vendors & Orders</h5>
						<a href="/admin/inventory-management.asp">Vendor dashboard</a><br/>
						<a href="/admin/PurchaseOrders.asp">Purchase orders</a><br/>
						<a href="/admin/add_company.asp">Manage vendors</a>

				</div>
				<div class="col">
					<h5>Products</h5>
						<a href="/admin/product-edit.asp?add-new-product=yes">Add new product</a><br/>
						<a href="/admin/old-products-sales.asp">Manage sales on old stock</a><br/>
						<a href="/admin/edit_select.asp?jewelry=save">Saved jewelry</a><br/>
						<a href="/admin/available-empty-bins.asp">Available empty bins</a>	
						<h5 class="mt-3">Research</h5>
						<a href="/admin/edits_logs.asp">Edit logs</a><br/>	
				</div>
				<div class="col">
					<h5>Etsy</h5>
						<a href="/admin/etsy-manage-inventory.asp?page=1">Etsy stock</a>		
				</div>
			</div>
		</div>
	</div>
</li>
<% end if  '=========== END PRE-ORDERS MENU ================================ %>

<% ' START packaging
If var_access_level = "Packaging" or var_access_level = "Admin" or var_access_level = "Manager" then  %>
	<li class="nav-item dropdown position-static border-right border-secondary">
		<a class="nav-link" href="#" id="PackagingDropdown"
				role="button" data-toggle="dropdown" aria-haspopup="true"
				aria-expanded="false">
				Packaging</a>
	<div class="dropdown-menu bg-dark w-100 m-0  border-0 rounded-0 pb-5"
				aria-labelledby="PackagingDropdown">
		<div class="container-fluid">
			<div class="row w-100">
				<div class="col">
					<h5>Invoices</h5>
						<a href="/admin/review-backorders.asp">Review backorders</a>
				</div>
				<div class="col">
					<h5>Import tracking #'s</h5>
						<a href="/admin/insertUPSTracking_numbers.asp">Import UPS #'s</a>
				</div>
				<div class="col">
					<h5>Labels</h5>
						<a href="/admin/update-labels.asp">Update labels</a><br/>
						<a href="/admin/batch-shipping.asp">Print Shipping Forms</a><br/><br/>
				</div>
				<div class="col">
		
				</div>
			</div>
		</div>
	</div>
</li>
<% end if ' packaging navigation %>


<% ' START Social Media
If var_access_level = "Inventory" and (user_name = "Charles" or var_access_level = "Admin") then  %>
	<li class="nav-item dropdown position-static border-right border-secondary">
		<a class="nav-link" href="#" id="SocialDropdown"
				role="button" data-toggle="dropdown" aria-haspopup="true"
				aria-expanded="false">
				Social Media</a>
	<div class="dropdown-menu bg-dark w-100 m-0  border-0 rounded-0 pb-5"
				aria-labelledby="SocialDropdown">
		<div class="container-fluid">
			<div class="row w-100">
				<div class="col">
						<a href="/admin/coupons_manage.asp">Manage coupons</a><br/>
						<a href="/admin/one-time-coupons.asp">One time use coupons</a><br/>
						<a href="/admin/secret-sale-items.asp">Secret sale items</a><br/>
						<a href="/admin/sliders/sliders.asp">Manage home page sliders</a><br/>
				</div>
			</div>
		</div>
	</div>
</li>
<% end if ' Social Media %>


<% ' START Photography menu
If var_access_level = "Photography" then  %>
	<li class="nav-item dropdown position-static border-right border-secondary">
		<a class="nav-link" href="#" id="PhotographyDropdown"
				role="button" data-toggle="dropdown" aria-haspopup="true"
				aria-expanded="false">
				Photography</a>
	<div class="dropdown-menu bg-dark w-100 m-0  border-0 rounded-0 pb-5"
				aria-labelledby="PhotographyDropdown">
		<div class="container-fluid">
			<div class="row w-100">
				<div class="col">
					<a href="/admin/product-edit.asp?add-new-product=yes">Add new product</a><br/>
					<a href="/admin/add_company.asp">Vendor list</a><br/>
					<a href="/admin/available-empty-bins.asp">Available empty bins</a><br/>
					<a href="/admin/thumbnails-review.asp">Manage thumbnails</a><br/>
					<a href="/admin/update-labels.asp">Update labels</a>
				</div>
				<div class="col">
					<h5>Site management</h5>
						<a href="/admin/sliders/sliders.asp">Manage home page sliders</a><br/>
				</div>
			</div>
		</div>
	</div>
</li>
<% end if ' END Photograhy menu %>


<% ' START Office Manager menu
If var_access_level = "Manager" then  %>
	<li class="nav-item dropdown position-static border-right border-secondary">
		<a class="nav-link" href="#" id="ManagerDropdown"
				role="button" data-toggle="dropdown" aria-haspopup="true"
				aria-expanded="false">
				Office manager</a>
	<div class="dropdown-menu bg-dark w-100 m-0 border-0 rounded-0 pb-5 "
				aria-labelledby="ManagerDropdown">
		<div class="container-fluid">
			<div class="row w-100">
				<div class="col">
					<h5>Site management</h5>
						<a href="/admin/coupons_manage.asp">Manage coupons</a><br/>
						<a href="/admin/one-time-coupons.asp">One time use coupons</a><br/>
						<a href="/admin/sliders/sliders.asp">Manage home page sliders</a><br/>
						<a href="/admin/manage_shippingmethods.asp">Shipping options</a>
				</div>
				<div class="col">
					<h5>Employee management</h5>
						<a href="/admin/edits_logs.asp">Edit logs</a><br/>
						<a href="/admin/packing-errors.asp">Packing error reports</a><br/>
						<a href="/admin/manage_users.asp">Manage admin users</a>
				</div>
				<div class="col">
					<h5>Temporary projects</h5>
						<a href="/admin/sandbox.asp">Enable sandbox testing</a><br/>
						<a href="/admin/thumbnails-review.asp">Manage thumbnails</a>
				</div>
				<div class="col">
		
				</div>
			</div>
		</div>
	</div>
</li>
<% end if ' END Office Manager menu %>


<% ' START admin only navigation
If var_access_level = "Admin" then  %>
	<li class="nav-item dropdown position-static border-right border-secondary">
		<a class="nav-link" href="#" id="AdminDropdown"
				role="button" data-toggle="dropdown" aria-haspopup="true"
				aria-expanded="false">
				Admin</a>
	<div class="dropdown-menu bg-dark w-100 m-0 border-0 rounded-0 pb-5"
				aria-labelledby="AdminDropdown">
		<div class="container-fluid">
			<div class="row w-100">
				<div class="col">
					<h5>Financials</h5>
						<a href="/admin/authnet-batches.asp">Batches</a>
				</div>
				<div class="col">
					<h5>Site management</h5>
					<a href="/admin/instagram.asp">Instagram orders</a><br/>	
					<a href="/admin/coupons_manage.asp">Manage coupons</a><br/>
					<a href="/admin/sliders/sliders.asp">Manage home page sliders</a><br/>
						<a href="/admin/one-time-coupons.asp">One time use coupons</a><br/>
						<a href="/admin/manage_shippingmethods.asp">Shipping options</a><br/>
						<a href="/admin/shipping-notice.asp">Shipping notice</a><br/>
						<a href="/admin/countries-manage.asp">Manage countries (front end & back end)</a><br/>
						<a href="/admin/seo-product-search-manager.asp">SEO Product search manager</a><br/>
						<a href="/admin/seo-title-description-tagging.asp">SEO Title & description tagging</a>
				</div>
				<div class="col">
					<h5>Employee management</h5>
						<a href="/admin/edits_logs.asp">Edit logs</a><br/>
						<a href="/admin/packing-errors.asp">Packing error reports</a><br/>
						<a href="/admin/manage_users.asp">Manage admin users</a>
			
				</div>
				<div class="col">
					<h5>Temporary projects</h5>				
						<a href="/admin/duplicate-customer-accounts.asp">Customer duplicate accounts</a><br/>
						<a href="/admin/sandbox.asp">Enable sandbox testing</a>
					<h5>Products</h5>
						<a href="/admin/edit_select.asp?jewelry=save">Saved jewelry</a>
				</div>
				<div class="col">
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
			</div>
		</div>
	</div>
</li>
<% end if ' Admin only navigation %>
</ul>
	<ul class="navbar-nav ml-auto">
<%
if var_access_level = "Admin" or var_access_level = "Manager" or user_name = "Nathan" or user_name = "Melissa" or user_name = "Anna" or user_name = "Rebekah" or user_name = "Sarena" then

'====== Get products that need to be reviewed count and use connection for it's page ======
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT title, date_added, added_by, picture, ProductID, active, reviewed_by_1, reviewed_by_2, review_date_1, review_date_2 FROM jewelry WHERE reviewed_by_2 IS NULL ORDER BY active DESC, date_added ASC"

set rsGetProductsToReview = Server.CreateObject("ADODB.Recordset")
rsGetProductsToReview.CursorLocation = 3 'adUseClient
rsGetProductsToReview.Open objCmd
total_products_to_review = rsGetProductsToReview.RecordCount

if total_products_to_review > 0 then
%>
<li class="nav-item"><a class="nav-link border-right border-secondary bg-primary" href="/admin/review-products.asp">Products (<%= total_products_to_review %>)</a></li>
<%
end if

SqlString = "SELECT COUNT(TBL_PhotoGallery.PhotoID) AS TotalPhotos FROM jewelry INNER JOIN TBL_PhotoGallery ON jewelry.ProductID = TBL_PhotoGallery.ProductID INNER JOIN ProductDetails ON TBL_PhotoGallery.DetailID = ProductDetails.ProductDetailID LEFT OUTER JOIN customers ON TBL_PhotoGallery.customerID = customers.customer_ID WHERE TBL_PhotoGallery.status = 0"
Set rsGetPhotoReviewCount = DataConn.Execute(SqlString)

if rsGetPhotoReviewCount.Fields.Item("TotalPhotos").Value > 0 then
%>
 <li class="nav-item"><a class="nav-link border-right border-secondary bg-warning" style="color:#000!important" href="/admin/approve-photos.asp">Photos (<%= rsGetPhotoReviewCount.Fields.Item("TotalPhotos").Value %>)</a></li>
<%
end if

SqlString = "SELECT  TOP (100) PERCENT COUNT(TBLReviews.ReviewID) AS TotalReviews FROM jewelry INNER JOIN TBLReviews ON jewelry.ProductID = TBLReviews.ProductID INNER JOIN ProductDetails ON TBLReviews.DetailID = ProductDetails.ProductDetailID INNER JOIN customers ON TBLReviews.customer_ID = customers.customer_ID WHERE (TBLReviews.status = N'pending')"
Set rsGetJewelryReviewCount = DataConn.Execute(SqlString)

if rsGetJewelryReviewCount.Fields.Item("TotalReviews").Value > 0 then
%>
      <li class="nav-item"><a class="nav-link border-right border-secondary bg-warning" style="color:#000!important" href="/admin/approve-reviews.asp">Reviews (<%= rsGetJewelryReviewCount.Fields.Item("TotalReviews").Value %>)</a></li>
<%
end if

end if ' access level for moderation
%>

<% If Not rsGetUser.EOF Or Not rsGetUser.BOF Then %>
	<li class="nav-item">
		<a class="btn btn-sm btn-secondary small mx-2 mt-1" href="/admin/admin_logout.asp">Logout <%=(rsGetUser.Fields.Item("name").Value)%></a>
	</li>

<% End If ' end Not rsGetUser.EOF Or NOT rsGetUser.BOF %>


<li class="nav-item"><!-- begin dark mode button -->
	<% if request.cookies("admindarkmode") <> "on" then
		darkchecked = "" 
	else 
		darkchecked = "checked"
	end if %>
	<style>
		.onoffswitch-inner:before {
			content: "LIGHT"
		}

		.onoffswitch-inner:after {
			content: "DARK"
		}
	</style>
	<span>
		<div class="onoffswitch">
			<input type="checkbox" name="onoffswitch" class="onoffswitch-checkbox" id="darkmode-switch" <%= darkchecked %>>
			<label class="onoffswitch-label" style="margin-bottom:0;margin-top:3px" for="darkmode-switch">
				<span class="onoffswitch-inner"></span>
				<span class="onoffswitch-switch"></span>
			</label>
		</div>
	</span>	
</li><!-- end dark mode button-->


			</ul>
</nav>


<script type="text/javascript" src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/bootstrap-v4.min.js"></script>
<script type="text/javascript" src="/js/js.cookie.js"></script>
<script language="JavaScript" type="text/JavaScript">

	// Dark mode toggle
	$("#darkmode-switch").on("click", function () {
	   console.log("test 3");
		if ($("#darkmode-switch").is(':checked')) {
			$('head').append('<link href="/CSS/baf-dark.min.css" rel="stylesheet" id="darkmode" />');
			$('link[rel=stylesheet][id="lightmode"]').prop('disabled', true);
			Cookies.set("admindarkmode", "on", { expires: 20*365});
			console.log("darkmode");
		} else {
			$('head').append('<link href="/CSS/baf.min.css" rel="stylesheet" id="lightmode" />');
			$('link[rel=stylesheet][id="darkmode"]').remove();
			Cookies.set("admindarkmode", "off", { expires: 20*365});
			console.log("lightmode");
		}
	});
	
	// Toggles
	$("#toggle_autoclave, #toggle_checkout_cards, #toggle_checkout_paypal").on("click", function () {
		$.ajax({
			method: "POST",
			url: "toggle.asp",
			data: {toggleItem: $(this).attr("id"), isChecked: $(this).is(":checked")}
		})		
	});	
</script>