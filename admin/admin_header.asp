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

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT InvoiceID FROM TBL_OrderSummary WHERE BackorderReview = 'Y'"

set rsBOReviews = Server.CreateObject("ADODB.Recordset")
rsBOReviews.CursorLocation = 3 'adUseClient
rsBOReviews.Open objCmd
total_backorders_to_review = rsBOReviews.RecordCount

set objcmd = Server.CreateObject("ADODB.command")
objcmd.ActiveConnection = DataConn
objcmd.CommandText = "SELECT DISTINCT TOP (100) PERCENT Count(*) AS total_to_anodize FROM dbo.TBL_OrderSummary AS ORS LEFT OUTER JOIN  dbo.sent_items AS SNT ON SNT.ID = ORS.InvoiceID AND ORS.item_price > 0 AND ORS.anodized_completed = 0 AND ORS.anodization_id_ordered > 0 WHERE (SNT.anodize = 1) ANd ship_code = 'paid'"
Set rsAnodizeCount = objcmd.Execute()

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT Count(*) AS Total_ToReview FROM sent_items WHERE shipped = 'Review'"
Set rsGetOrdersToReview = objCmd.Execute()

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM TBL_Barcodes_SortOrder" 
Set rs_getsections = objCmd.Execute()
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
If var_access_level = "Admin" or var_access_level = "Manager" or var_access_level = "Customer service" or user_name = "Melissa" then
%>
<li class="nav-item border-right border-secondary">
	<a class="nav-link" href="/admin/landing/customer-service.asp">Customer Service</i>
	</a>
</li>
<% end if '=========== END CUSTOMER SERVICE MENU ================================ %>

<%
'=========== START PRODUCT MANAGEMENT MENU ========================================
If var_access_level = "Admin" or var_access_level = "Manager" or var_access_level = "Inventory" or var_access_level = "Photography" then  
%>
<li class="nav-item border-right border-secondary">
	<a class="nav-link" href="/admin/landing/product-management.asp">Product Management</i>
	</a>
</li>
<% end if '=========== END PRODUCT MANAGEMENT MENU ================================ %>

<% If var_access_level = "Packaging" or var_access_level = "Admin" or var_access_level = "Manager" then  %>
<li class="nav-item border-right border-secondary">
	<a class="nav-link" href="/admin/landing/packaging.asp">Packaging</i>
	</a>
</li>
<% end if  %>

<% If var_access_level = "Social Media" then  %>
<li class="nav-item border-right border-secondary">
	<a class="nav-link" href="/admin/landing/social-media.asp">Social Media</i>
	</a>
</li>
<% end if %>

<% If var_access_level = "Manager" or var_access_level = "Admin" then %>
<li class="nav-item border-right border-secondary">
	<a class="nav-link" href="/admin/landing/management.asp">Management</i>
	</a>
</li>
<% end if %>
</ul>
	<ul class="navbar-nav ml-auto">
	<%  
	If user_name <> "Rebekah" and total_backorders_to_review > 0  then  %>
	<li class="nav-item"><span class="nav-link text-light border-right border-light">Backorders<span class="badge badge-warning ml-2"><%= total_backorders_to_review %></span> </span></li>
	<% end if  %>	

<% if  user_name = "Adrienne" or user_name = "Melissa" or user_name = "Andres" or  user_name = "Amanda" or  user_name = "Ellen"    then %>
<li class="nav-item"><a class="nav-link border-right border-light bg-secondary" href="/admin/anodization-orders.asp">Anodization<span class="badge badge-warning ml-2"><%= rsAnodizeCount("total_to_anodize") %></span></a></li>
<% end if %>
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
<li class="nav-item"><a class="nav-link border-right border-light bg-secondary" href="/admin/review-products.asp">Products<span class="badge badge-warning ml-2"><%= total_products_to_review %></span></a></li>
<%
end if

SqlString = "SELECT COUNT(TBL_PhotoGallery.PhotoID) AS TotalPhotos FROM jewelry INNER JOIN TBL_PhotoGallery ON jewelry.ProductID = TBL_PhotoGallery.ProductID INNER JOIN ProductDetails ON TBL_PhotoGallery.DetailID = ProductDetails.ProductDetailID LEFT OUTER JOIN customers ON TBL_PhotoGallery.customerID = customers.customer_ID WHERE TBL_PhotoGallery.status = 0"
Set rsGetPhotoReviewCount = DataConn.Execute(SqlString)

if rsGetPhotoReviewCount.Fields.Item("TotalPhotos").Value > 0 then
%>
 <li class="nav-item"><a class="nav-link border-right border-light bg-secondary" href="/admin/approve-photos.asp">Photos<span class="badge badge-warning ml-2"><%= rsGetPhotoReviewCount.Fields.Item("TotalPhotos").Value %></span></a></li>
<%
end if

SqlString = "SELECT  TOP (100) PERCENT COUNT(TBLReviews.ReviewID) AS TotalReviews FROM jewelry INNER JOIN TBLReviews ON jewelry.ProductID = TBLReviews.ProductID INNER JOIN ProductDetails ON TBLReviews.DetailID = ProductDetails.ProductDetailID INNER JOIN customers ON TBLReviews.customer_ID = customers.customer_ID WHERE (TBLReviews.status = N'pending')"
Set rsGetJewelryReviewCount = DataConn.Execute(SqlString)

if rsGetJewelryReviewCount.Fields.Item("TotalReviews").Value > 0 then
%>
      <li class="nav-item"><a class="nav-link border-right border-light bg-secondary" href="/admin/approve-reviews.asp">Reviews<span class="badge badge-warning ml-2"><%= rsGetJewelryReviewCount.Fields.Item("TotalReviews").Value %></span></a></li>
<%
end if

end if ' access level for moderation
%>

<% If Not rsGetUser.EOF Or Not rsGetUser.BOF Then %>
	<li class="nav-item">
		<a class="btn btn-sm btn-secondary small mx-2 mt-1" href="/admin/admin_logout.asp">Logout <%=(rsGetUser.Fields.Item("name").Value)%></a>
	</li>

<% End If ' end Not rsGetUser.EOF Or NOT rsGetUser.BOF %>
<% If rsGetUser.EOF And rsGetUser.BOF Then %>
	<li class="nav-item">
		<a class="btn btn-sm btn-secondary small mx-2 mt-1" href="/admin/login.asp?login=yes">Login</a>
	</li>
	<% 
	End If ' end rsGetUser.EOF And rsGetUser.BOF 
	%>

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
	
</script>