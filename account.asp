<% @LANGUAGE="VBSCRIPT" %>
<%
	page_title = "Order History"
	page_description = "Your Bodyartforms order history."
	page_keywords = ""
	
' Clear temporary account if admin is viewing
if request.querystring("cleartemp") = "yes" then
	response.cookies("flag-tempid") = ""
	session("admin_tempcustid") = ""
end if
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<%
var_flagged = ""

' Pull the customer information from a cookie
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM customers  WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
		Set rsGetUser = objCmd.Execute()
		
		
If Not rsGetUser.EOF Or Not rsGetUser.BOF Then ' Only run this info if a match was found

	if rsGetUser.Fields.Item("Flagged").Value = "Y" then
		var_flagged = "yes"
	end if
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT shipped, email FROM sent_items WHERE email = ? AND (shipped = 'Flagged' OR shipped = 'Chargeback')"
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,250,rsGetUser.Fields.Item("email").Value))
	set rsGetFlaggedOrders = objCmd.Execute()
	
	if NOT rsGetFlaggedOrders.eof then
		var_flagged = "yes"
	end if

	'Get country list for drop downs
	Set rsGetCountrySelect = Server.CreateObject("ADODB.Recordset")
	rsGetCountrySelect.ActiveConnection = DataConn
	rsGetCountrySelect.Source = "SELECT * FROM dbo.TBL_Countries WHERE Display = 1 ORDER BY Country ASC "
	rsGetCountrySelect.CursorLocation = 3 'adUseClient
	rsGetCountrySelect.LockType = 1 'Read-only records
	rsGetCountrySelect.Open()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM sent_items WHERE ship_code = 'paid' AND customer_ID = ? ORDER BY ID DESC"
	'  AND date_order_placed > DATEADD(year, -3, GETDATE())'
	objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,rsGetUser.Fields.Item("customer_ID").Value))
	
	set rsGetOrders = Server.CreateObject("ADODB.Recordset")
	rsGetOrders.CursorLocation = 3 'adUseClient
	rsGetOrders.Open objCmd
	total_records = rsGetOrders.RecordCount
	rsGetOrders.PageSize = 10
	intPageCount = rsGetOrders.PageCount
	
	' Variables for paging
	Select Case Request("Action")
		case "<<"
			intpage = 1
		case "<"
			intpage = Request("intpage")-1
			if intpage < 1 then intpage = 1
		case ">"
			intpage = Request("intpage")+1
			if intpage > intPageCount then intpage = IntPageCount
		Case ">>"
			intpage = intPageCount
		case else
			intpage = 1
	end select	
		
		
End if ' Only run this info if a match was found

%>

<div class="display-5">
		Order History
	</div>
	<a class="text-info pointer" data-toggle="modal" data-target="#LegendModal"><i class="fa fa-question-circle"></i> Submitting reviews &amp; photos</a>
<%
if session("admin_tempcustid") <> "" then %>
	<div class="alert alert-success">Admin viewing
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<a href="account.asp?cleartemp=yes">Reset</a>
	</div>
<% end if %>
<!--#include virtual="/accounts/inc-account-navigation.asp" -->
<% If rsGetUser.EOF or var_flagged = "yes" Then
%>
	<div class="alert alert-danger">Not logged in or no account found
		<br/><br/>
		<a class="btn btn-outline-secondary" data-toggle="modal" data-target="#signin" href="#">Click here to sign in</a>

	</div>
<% else %>

<% If rsGetOrders.eof Then %>
	<h6>
		No orders found
	</h6>
<% else %>

<!--#include virtual="/accounts/inc-orders-paging.asp" -->

<% 

rsGetOrders.AbsolutePage = intPage '======== PAGING
For intRecord = 1 To rsGetOrders.PageSize

' Declare variables during loop

	paymethod = rsGetOrders.Fields.Item("pay_method").Value
	varstatus = rsGetOrders.Fields.Item("shipped").Value
	
	tracking_num = ""
	tracking_type = ""
	if instr(rsGetOrders.Fields.Item("shipping_type").Value, "DHL") > 0 then
		tracking_num = rsGetOrders.Fields.Item("USPS_tracking").Value
		tracking_type = "dhl"
	else
		tracking_num = rsGetOrders.Fields.Item("USPS_tracking").Value
		tracking_type = "usps"
	end if
	if instr(rsGetOrders.Fields.Item("shipping_type").Value, "UPS") then
		tracking_num = rsGetOrders.Fields.Item("UPS_tracking").Value
		tracking_type = "ups"
	end if
	
	' variable for date shipped
	var_date_shipped = ""
	if rsGetOrders.Fields.Item("date_sent").Value <> "" then
		var_date_shipped = MonthName(Month(rsGetOrders.Fields.Item("date_sent").Value),1) & " " & Day(rsGetOrders.Fields.Item("date_sent").Value) & ", " & Year(rsGetOrders.Fields.Item("date_sent").Value)
	end if	
	
	' variable for order status
	var_order_status = ""
	if rsGetOrders.Fields.Item("shipped").Value = "Pending..." then
		var_order_status = "Pending shipment"
	elseif rsGetOrders.Fields.Item("shipped").Value = "Review" then
		var_order_status = "Pending shipment"
	elseif rsGetOrders.Fields.Item("shipped").Value = "Shipped" then
		var_order_status = "Shipped on " & var_date_shipped
	else
		var_order_status = rsGetOrders.Fields.Item("shipped").Value
	end if

	'pre-order status variables
	if rsGetOrders.Fields.Item("shipped").Value = "PRE-ORDER REVIEW" then
		var_order_status = "Your pre-order items are currently under review for approval before being submitted to the manufacturer to be made."
	end if
	if rsGetOrders.Fields.Item("shipped").Value = "PRE-ORDER APPROVED" then
		var_order_status = "Your pre-order items have been reviewed and approved and will soon be sent to the manufacturer to be made for you."
	end if
	if rsGetOrders.Fields.Item("shipped").Value = "ON ORDER" then
		var_order_status = "Your pre-order items have been sent to the manufacturer to be made and at this time we are waiting for them to arrive. We'll ship them to you as soon as they come in!"
	end if	
	
	' variable for ship time ETA
	var_ship_eta = ""
	if rsGetOrders.Fields.Item("ship_code").Value = "paid" AND rsGetOrders.Fields.Item("shipped").Value = "Pending..." then

		If WeekDayName(WeekDay(date())) = "Saturday" OR WeekDayName(WeekDay(date())) = "Sunday" then
			var_ship_eta = " on Monday"
		end if
		
		If Time() > "08:00:00 AM" AND WeekDayName(WeekDay(date())) <> "Saturday" AND WeekDayName(WeekDay(date())) <> "Sunday" AND WeekDayName(WeekDay(date())) <> "Friday" then
			var_ship_eta = " tomorrow"
		end if
		
		If Time() > "08:00:00 AM" AND WeekDayName(WeekDay(date())) = "Friday" then
			var_ship_eta = " on Monday"
		end if
		
		If Time() < "08:00:00 AM" AND WeekDayName(WeekDay(date())) <> "Saturday" AND WeekDayName(WeekDay(date())) <> "Sunday" then
			var_ship_eta = " today"
		end if

	end if
	
	' variable that allows changes to order
	var_allow_changes = ""
	if (paymethod = "Visa" OR paymethod = "Mastercard" OR paymethod = "MasterCard" OR paymethod = "Discover" OR paymethod = "American Express") AND (varstatus = "Pending..." OR varstatus = "Review" OR varstatus = "ON HOLD") then
		var_allow_changes = "yes"
	end if
	
	' variable to allow order to report a problem
	var_allow_report_problem = ""
	if (date() <= rsGetOrders.Fields.Item("date_sent").Value + 30 AND date() >= rsGetOrders.Fields.Item("date_sent").Value) AND varstatus = "Shipped"  then
		var_allow_report_problem = "yes"
	end if
	
	
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT OrderDetailID FROM TBL_OrderSummary  WHERE ProductID = 2424 AND InvoiceID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,rsGetOrders.Fields.Item("ID").Value))
Set rsContainsGiftCert = objCmd.Execute()

' Find out whether a gift cert was purchased in the order so that the user can't cancel the order via the website
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT sent_items.total_store_credit as store_credit, sent_items.total_gift_cert as gift_cert,  sent_items.shipping_rate - sent_items.total_preferred_discount - sent_items.total_coupon_discount - sent_items.total_free_credits + sent_items.total_sales_tax - sent_items.total_store_credit - sent_items.total_gift_cert - sent_items.total_returns AS total_discount_taxes, SUM(TBL_OrderSummary.qty * TBL_OrderSummary.item_price) AS subtotal FROM sent_items INNER JOIN TBL_OrderSummary ON sent_items.ID = TBL_OrderSummary.InvoiceID WHERE (sent_items.ID = ?) GROUP BY  sent_items.total_store_credit, sent_items.total_gift_cert,  sent_items.shipping_rate - sent_items.total_preferred_discount - sent_items.total_coupon_discount - sent_items.total_free_credits + sent_items.total_sales_tax - sent_items.total_store_credit - sent_items.total_gift_cert - sent_items.total_returns"
objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,rsGetOrders.Fields.Item("ID").Value))
Set rsGetOrderTotals = objCmd.Execute()
%>

<div class="card border-secondary my-5">
		<div class="card-header px-0 py-1">
			<div class="container-fluid">
				<div class="row">
					<div class="col-6 col-md-4 small">
						<div class="font-weight-bold text-secondary">ORDER STATUS</div>
						<div class="text-secondary" id="status-<%= rsGetOrders.Fields.Item("ID").Value %>">
							<%= var_order_status %><%= var_ship_eta %>
						</div>
					</div>
					<div class="col-6 col-md-4 small">
							<div class="font-weight-bold text-secondary">PLACED ON</div>
							<div class="text-secondary">
									<%= MonthName(Month(rsGetOrders.Fields.Item("date_order_placed").Value),1)%>&nbsp;<%= Day(rsGetOrders.Fields.Item("date_order_placed").Value)%>, <%= Year(rsGetOrders.Fields.Item("date_order_placed").Value)%>
							
						</div>
					</div>
					<div class="col-6 col-md-4 small">
							<div class="font-weight-bold text-secondary">
						Invoice # <span class="d-lg-none">&zwj;<%= rsGetOrders.Fields.Item("ID").Value %></span><span class="d-none d-lg-inline-block"><%= rsGetOrders.Fields.Item("ID").Value %></span></div><!-- odd character to force mobile phones to not view invoice # as a mobile phone number -->
						<div>
							<a href="" class="dropdown-toggle text-secondary" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
							  View shipping address
							</a>
							<div class="dropdown-menu px-3" id="address-box-<%= rsGetOrders.Fields.Item("ID").Value %>">
								<% if (rsGetOrders.Fields.Item("company").Value) <> "" then %>
								<%=(rsGetOrders.Fields.Item("company").Value)%><br>
								<% end if %>
								<%=(rsGetOrders.Fields.Item("customer_first").Value)%> &nbsp;<%=(rsGetOrders.Fields.Item("customer_last").Value)%><br>
								<%=(rsGetOrders.Fields.Item("address").Value)%> <br>
								<% if (rsGetOrders.Fields.Item("address2").Value) <> "" then %>
								<%=(rsGetOrders.Fields.Item("address2").Value)%> <br>
								<% end if %>
								<%=(rsGetOrders.Fields.Item("city").Value)%>, <%=(rsGetOrders.Fields.Item("state").Value)%><%=(rsGetOrders.Fields.Item("province").Value)%>&nbsp;&nbsp;<%=(rsGetOrders.Fields.Item("zip").Value)%><br>
								<%=(rsGetOrders.Fields.Item("country").Value)%>
							  </div>
						</div>
				</div>
			</div>
			</div>
		</div>
		<div class="card-body">
		 
	
		
	  


<% if not rsGetOrderTotals.eof then %>

	<% if tracking_num <> "" then
	if tracking_type = "usps" then %>
		<button class="btn btn-purple btn-sm my-1 track" type="button" data-num="<%= tracking_num %>" data-invoice="<%= rsGetOrders.Fields.Item("ID").Value %>" data-url="/usps_tools/usps_tracking.asp?id=" data-toggle="modal" data-target="#trackModal">Track package</button>
	<%	elseif tracking_type = "dhl" then %>
		<button class="btn btn-purple btn-sm my-1 track" type="button" data-num="<%= tracking_num %>" data-invoice="<%= rsGetOrders.Fields.Item("ID").Value %>" data-url="/dhl/dhl-tracking.asp?tracking=" data-toggle="modal" data-target="#trackModal">Track package</button>
	<%
	else ' if ups%>
		UPS tracking # <%= tracking_num %>
	<%
	end if 'if tracking type usps or ups
	end if ' tracking_num <> ""
	%>
<%
paid_status = rsGetOrders.Fields.Item("ship_code").Value
varstatus = rsGetOrders.Fields.Item("shipped").Value
orderscanned = rsGetOrders.Fields.Item("ScanInvoice_Timestamp").Value

if paid_status = "paid" and isnull(orderscanned) and rsGetOrders.Fields.Item("pay_method").Value <> "Afterpay" and (varstatus = "Pending..." OR varstatus = "Pending shipment" OR varstatus = "Review" OR varstatus = "PRE-ORDER REVIEW") then
%>
	<button class="btn btn-purple btn-sm my-1 btn-addon-modal" id="addon-<%= rsGetOrders.Fields.Item("ID").Value %>" type="button" data-invoice="<%= rsGetOrders.Fields.Item("ID").Value %>" data-toggle="modal" data-target="#AddonModal">Add item(s) to order</button>
<% 
end if

if paid_status = "paid" and rsGetOrderTotals.Fields.Item("store_credit").Value = 0 and rsGetOrderTotals.Fields.Item("gift_cert").Value = 0 and (varstatus = "Pending..." OR varstatus = "Pending shipment" OR varstatus = "Review" OR varstatus = "PRE-ORDER REVIEW") then

if rsContainsGiftCert.BOF and rsContainsGiftCert.EOF then
%>
	<button class="btn btn-purple btn-sm my-1 btn-cancel-modal" id="cancel-<%= rsGetOrders.Fields.Item("ID").Value %>" type="button" data-invoice="<%= rsGetOrders.Fields.Item("ID").Value %>" data-toggle="modal" data-target="#cancelOrderModal">Cancel order</button>
	
	<button class="btn btn-purple btn-sm my-1 btn-update-address-modal" type="button" data-invoice="<%= rsGetOrders.Fields.Item("ID").Value %>"  data-country="<%= rsGetOrders.Fields.Item("country").Value %>" data-toggle="modal" data-target="#updateAddressModal">Update shipping address</button>
<% 
end if	' rsContainsGiftCert.EOF
end if

 ' -------------- Only display survey link if order is $10 or more ------------- 
If rsGetOrderTotals.Fields.Item("subtotal").Value + rsGetOrderTotals.Fields.Item("total_discount_taxes").Value >= 10 then 

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT InvoiceID FROM TBL_Surveys WHERE InvoiceID = ?"
		objCmd.Prepared = true
		objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,rsGetOrders.Fields.Item("ID").Value))
		Set rsGetSurvey = objCmd.Execute()
		
	
	If rsGetSurvey.EOF then 
		If rsGetOrders.Fields.Item("shipped").Value = "Shipped" AND now() >= (rsGetOrders.Fields.Item("date_sent").Value + 14) AND rsGetOrders.Fields.Item("date_sent").Value >= now() - 730 then
		
		
		
	Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")

	password = "3uBRUbrat77V"
	data = rsGetOrders.Fields.Item("ID").Value
	invoice_encrypted = objCrypt.Encrypt(password, data)
	Set objCrypt = Nothing
	%>
<button class="btn btn-purple btn-sm my-1 btn-survey-modal btn-open-survey-<%= rsGetOrders.Fields.Item("ID").Value %>" type="button" data-id="<%= invoice_encrypted %>" data-toggle="modal" data-target="#submitSurveyModal">
	Take order survey (.50¢ store credit)</button>
	<% 
		End if
	End if %>
<% End if  ' Display survey if order is $10 or more
%>	
	
	
	<% if var_allow_changes = "yes" then %>
	
	<% end if
	if var_allow_report_problem = "yes" then %>
	<button class="btn btn-purple btn-sm my-1 btn-report-problem" type="button" data-invoice="<%= rsGetOrders.Fields.Item("ID").Value %>" data-toggle="modal" data-target="#ReportProblemModal">Report order problem</button>
	<button class="btn btn-purple btn-sm my-1" type="button" data-toggle="modal" data-target="#ReturnPolicyModal">Returns</button>
	<% end if %> 

	<div class="d-flex flex-row flex-wrap">
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TOP (100) PERCENT TBL_OrderSummary.InvoiceID AS ID, TBL_OrderSummary.OrderDetailID, TBL_OrderSummary.qty, jewelry.title, ProductDetails.ProductDetail1, TBL_OrderSummary.returned, TBL_OrderSummary.item_price, TBL_OrderSummary.backorder, TBL_OrderSummary.ProductReviewed, TBL_OrderSummary.ProductPhotographed, TBL_OrderSummary.PreOrder_Desc, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.free, jewelry.picture, TBL_PhotoGallery.status, TBLReviews.status AS ReviewStatus, TBLReviews.review_rating, jewelry.jewelry, TBL_PhotoGallery.PhotoID, ISNULL(TBL_OrderSummary.notes,'') as notes, ProductDetails.ProductDetailID, ProductDetails.ProductID, title + ' ' + Gauge + ' ' + Length + ' ' + ProductDetail1 as 'concat-title' FROM  dbo.TBL_OrderSummary INNER JOIN dbo.ProductDetails ON dbo.TBL_OrderSummary.DetailID = dbo.ProductDetails.ProductDetailID INNER JOIN dbo.jewelry ON dbo.ProductDetails.ProductID = dbo.jewelry.ProductID FULL OUTER JOIN dbo.TBLReviews ON dbo.TBL_OrderSummary.OrderDetailID = dbo.TBLReviews.ReviewOrderDetailID FULL OUTER JOIN dbo.TBL_PhotoGallery ON dbo.TBL_OrderSummary.OrderDetailID = dbo.TBL_PhotoGallery.OrderDetailID WHERE TBL_OrderSummary.InvoiceID = ? ORDER BY OrderDetailID ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,rsGetOrders.Fields.Item("ID").Value))
Set rsGetOrderDetails = objCmd.Execute()

var_free_items = ""
While not rsGetOrderDetails.eof
	
	var_order_detailid = rsGetOrderDetails.Fields.Item("OrderDetailID").Value
	
	var_bo_status = ""
	var_bo_text = ""
	'set backorder variable
	if rsGetOrderDetails.Fields.Item("backorder").Value = 1 then
		var_bo_status = "backorder-on"
		var_bo_text = "<div class=""badge badge-warning my-1 d-block rounded-0"">ON BACKORDER</div>"
	end if 
	
	' ONLY ALLOW REVIEWS AND PHOTOS TO BE SUBMITTED UP TO 2 YEARS AFTER THE ORDER WAS PLACED
	' variable for write a review and photos status
	var_review_info = ""
	var_photo_info = ""
	if  rsGetOrders.Fields.Item("date_order_placed").Value > DateAdd("yyyy",-2,date()) AND rsGetOrderDetails.Fields.Item("returned").Value <> 1 AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 1430 AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 530 AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 3928 AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 15385 AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 3611 AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 3587  AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 3086  AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 3704 AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 1649 AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 4287 AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 1483 AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 3926 AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 3803 AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 1851 AND rsGetOrderDetails.Fields.Item("jewelry").Value <> "save" then

	'    JEWELRY REVIEWS	
	if rsGetOrderDetails.Fields.Item("ProductReviewed").Value = "N" then
		var_review_info = "<button class=""btn btn-light btn-sm btn-update-attributes btn-clear-forms btn-review-modal review-phase1-" & rsGetOrderDetails.Fields.Item("OrderDetailID").Value & """ type=""button"" data-orderitemid=""" & rsGetOrderDetails.Fields.Item("OrderDetailID").Value &""" data-title=""" & rsGetOrderDetails.Fields.Item("concat-title").Value &""" data-toggle=""modal"" data-target=""#reviewModal""><i class=""fa fa-star-half-o fa-lg""></i></button>"	
	end if
	if rsGetOrderDetails.Fields.Item("ReviewStatus").Value = "pending" then
		var_review_info = "<span class=""btn btn-sm alert-success""><i class=""fa fa-star-half-o fa-lg""></i></span>"
	end if
	if rsGetOrderDetails.Fields.Item("ReviewStatus").Value = "rejected" then
			var_review_info = "<button class=""btn btn-danger btn-sm btn-update-attributes btn-clear-forms btn-review-modal review-phase1-" & rsGetOrderDetails.Fields.Item("OrderDetailID").Value & """ type=""button"" data-orderitemid=""" & rsGetOrderDetails.Fields.Item("OrderDetailID").Value &""" data-title=""" & rsGetOrderDetails.Fields.Item("concat-title").Value &""" data-toggle=""modal"" data-target=""#reviewModal""><i class=""fa fa-star-half-o fa-lg""></i></button>"
	end if ' for review
	if rsGetOrderDetails.Fields.Item("ReviewStatus").Value = "accepted" then
		var_review_info = ""
	end if
	
	'    PHOTO REVIEWS
	if rsGetOrderDetails.Fields.Item("ProductPhotographed").Value <> "Y" then
		var_photo_info = "<button class=""btn btn-light btn-sm ml-2 btn-update-attributes btn-clear-forms btn-photo-modal photo-phase1-" & rsGetOrderDetails.Fields.Item("OrderDetailID").Value & """ type=""button"" data-productid=""" & rsGetOrderDetails.Fields.Item("ProductID").Value &""" data-detailid=""" & rsGetOrderDetails.Fields.Item("ProductDetailID").Value &""" data-orderitemid=""" & rsGetOrderDetails.Fields.Item("OrderDetailID").Value &""" data-title=""" & rsGetOrderDetails.Fields.Item("concat-title").Value &""" data-toggle=""modal"" data-target=""#submitPhotoModal""><i class=""fa fa-camera fa-lg""></i></button>"
	end if
	if rsGetOrderDetails.Fields.Item("status").Value = 0 AND rsGetOrderDetails.Fields.Item("ProductPhotographed").Value = "Y" then
		var_photo_info = "<span class=""btn btn-sm alert-success ml-2 ""><i class=""fa fa-camera fa-lg""></i></span>"
	end if

	if isnull(rsGetOrderDetails.Fields.Item("status").Value) AND rsGetOrderDetails.Fields.Item("ProductPhotographed").Value = "Y" then
		var_photo_info = "<button class=""btn btn-danger btn-sm ml-2 btn-update-attributes btn-clear-forms btn-photo-modal photo-phase1-" & rsGetOrderDetails.Fields.Item("OrderDetailID").Value & """ type=""button"" data-productid=""" & rsGetOrderDetails.Fields.Item("ProductID").Value &""" data-detailid=""" & rsGetOrderDetails.Fields.Item("ProductDetailID").Value &""" data-orderitemid=""" & rsGetOrderDetails.Fields.Item("OrderDetailID").Value &""" data-title=""" & rsGetOrderDetails.Fields.Item("concat-title").Value &""" data-toggle=""modal"" data-target=""#submitPhotoModal""><i class=""fa fa-camera fa-lg""></i></button>"
	end if
	
	end if ' does not match all the products not allowed for reviews and photos

	' if it's a gauge card or stickers then don't show them in full, just show them as "addon free items"

	if rsGetOrderDetails.Fields.Item("notes").Value = "FREE" then
	var_free_items = var_free_items & "FREE: <span class=""mr-2"">" & rsGetOrderDetails.Fields.Item("qty").Value & "</span> " & rsGetOrderDetails.Fields.Item("gauge").Value & " " & rsGetOrderDetails.Fields.Item("length").Value & " " &   rsGetOrderDetails.Fields.Item("ProductDetail1").Value &  " " & rsGetOrderDetails.Fields.Item("title").Value & "<br/>"
	end if

	if instr(rsGetOrderDetails.Fields.Item("notes").Value, "FREE") <= 0 then
%>
<div class="col-6 col-sm-4 col-md-4 col-xl-3 col-lg-4 my-3 px-1 px-md-2 text-center order-history">	
	<div class="container-fluid p-0 m-0">
		<div class="row p-0 m-0">
			<div class="col p-0 m-0">

					
			
				
				<div class="position-relative">
						<a href="/productdetails.asp?ProductID=<%= rsGetOrderDetails.Fields.Item("ProductID").Value %>" target="_self">
				<img class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetOrderDetails.Fields.Item("picture").Value %>" alt="Product photo"></a>
				<div class="position-absolute order-img-icons p-0 m-0">
						<%= var_review_info %>
						<%= var_photo_info %>
					</div>
			</div>
			<% if rsGetOrderDetails.Fields.Item("PreOrder_Desc").Value <> "" AND rsGetOrderDetails.Fields.Item("ProductID").Value <> 2424 then %>
				<button class="btn btn-sm btn-outline-secondary btn-block my-1 btn-moreInfo-modal" type="button" data-toggle="modal" data-target="#MoreInfoModal" data-preorderSpecs="<%= Server.HTMLEncode(rsGetOrderDetails.Fields.Item("PreOrder_Desc").Value) %>">Pre-Order Details</button>
			<% end if %>
			<%= var_bo_text %>
			<div class="small font-weight-bold"><span class="pr-3">Qty <%= rsGetOrderDetails.Fields.Item("qty").Value %></span><%=FormatCurrency((rsGetOrderDetails.Fields.Item("item_price").Value)*(rsGetOrderDetails.Fields.Item("qty").Value),2)%></div>
			<div class="small">
			<%=(rsGetOrderDetails.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetOrderDetails.Fields.Item("Length").Value)%>&nbsp;<%=(rsGetOrderDetails.Fields.Item("ProductDetail1").Value)%>&nbsp;<%=(rsGetOrderDetails.Fields.Item("title").Value)%>
		</div>
			
			
		</div><!-- col -->
	</div><!-- row -->
		</div>	<!-- container-fluid end --> 
	</div><!-- flex column -->
<%
end if ' don't display free items
rsGetOrderDetails.Movenext()

wend
%>
</div><!-- end flex row -->
<% if var_free_items <> "" then %>
	<div class="small">
		<%= var_free_items %>
	</div>	
<% end if %>
<% if rsGetOrders.Fields.Item("item_description").Value <> "" or rsGetOrders.Fields.Item("customer_comments").Value <> "" then 

var_item_description = replace(rsGetOrders.Fields.Item("item_description").Value,"<br/>","")
var_item_description = replace(var_item_description,"ORDER UPDATED","")
var_item_description = replace(var_item_description,"ADDRESS UPDATED","")
var_item_description = replace(var_item_description,"SHIPPING METHOD UPDATED","")
%>
<% If rsGetOrders.Fields.Item("date_sent").Value >= now() - 182 then ' only show comments if less than 6 months old %>	
			<%= var_item_description %>
	<% end if %>	
<% end if %>

	
			
		
</div><!-- end card body -->
<div class="card-footer small">
	Subtotal <%= FormatCurrency(rsGetOrderTotals.Fields.Item("subtotal").Value, -1, -2, -0, -2)%><br/>	
<%

' Array for invoice totals
ReDim arrTotals(2,6) 

'arrTotals(col,row)
arrTotals(0,0) = "10% preferred discount" 
arrTotals(1,0) = "total_preferred_discount" 
total_preferred_discount = rsGetOrders.Fields.Item("total_preferred_discount").Value
arrTotals(2,0) = "&#8722;"
arrTotals(0,1) = "Coupon discount" 
arrTotals(1,1) = "total_coupon_discount" 
total_coupon_discount = rsGetOrders.Fields.Item("total_coupon_discount").Value
arrTotals(2,1) = "&#8722;" 
arrTotals(0,2) = "Tax" 
arrTotals(1,2) = "total_sales_tax" 
total_sales_tax = rsGetOrders.Fields.Item("total_sales_tax").Value
arrTotals(2,2) = "&nbsp;&nbsp;"
arrTotals(0,3) = "Gift certificate" 
arrTotals(1,3) = "total_gift_cert"
total_gift_cert = rsGetOrders.Fields.Item("total_gift_cert").Value 
arrTotals(2,3) = "&#8722;"
arrTotals(0,4) = "Free gift (USE NOW) credits" 
arrTotals(1,4) = "total_free_credits" 
total_free_credits = rsGetOrders.Fields.Item("total_free_credits").Value
arrTotals(2,4) = "&#8722;"
arrTotals(0,5) = "Store account credit" 
arrTotals(1,5) = "total_store_credit"
total_store_credit = rsGetOrders.Fields.Item("total_store_credit").Value
arrTotals(2,5) = "&#8722;"
arrTotals(0,6) = "Returned item(s)" 
arrTotals(1,6) = "total_returns"
total_returns = rsGetOrders.Fields.Item("total_returns").Value
arrTotals(2,6) = "&#8722;"


For i = 0 to UBound(arrTotals, 2) 

	if rsGetOrders.Fields.Item(arrTotals(1,i)).Value <> 0 then
%>
			<%= arrTotals(0,i) %>&nbsp;<%= arrTotals(2,i) %><%= FormatCurrency(rsGetOrders.Fields.Item(arrTotals(1,i)).Value, -1, -2, -0, -2) %><br/>

<% 
	end if ' if i > 2 or values not 0
next ' loop through totals array
%>


			<% if rsGetOrders.Fields.Item("shipping_type").Value <> "" then %>
			<%= replace(replace(replace(replace(rsGetOrders.Fields.Item("shipping_type").Value, "4) ", ""), "3) ", ""), "2) ", ""), "1) ", "") %>
			<% end if %>&nbsp;
			<%= FormatCurrency(rsGetOrders.Fields.Item("shipping_rate").Value, -1, -2, -0, -2) %>
			<div class="font-weight-bold">
			TOTAL <% if InvoiceTotal < 0 then %>0<% else %><%= FormatCurrency(rsGetOrderTotals.Fields.Item("subtotal").Value + rsGetOrderTotals.Fields.Item("total_discount_taxes").Value, -1, -2, -0, -2)  %><% end if %>
		</div>

<% end if %><!-- if not rsGetOrderTotals.eof then -->
</div><!-- end card footer -->
</div><!-- end card block -->               
<% 
rsGetOrders.MoveNext()
If rsGetOrders.EOF Then Exit For  ' ====== PAGING
Next ' ====== PAGING
%>

<!--#include virtual="/accounts/inc-orders-paging.asp" -->       

<% end if '   rsOrders.eof 
%>  

<!-- Track order modal -->
<div class="modal fade" id="trackModal" tabindex="-1" role="dialog" aria-labelledby="headTracking" aria-hidden="true">
		<div class="modal-dialog" role="document">
		  <div class="modal-content">
			<div class="modal-header">
			  <h5 class="modal-title" id="headTracking">Order Tracking History</h5>
			  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
				<span aria-hidden="true">&times;</span>
			  </button>
			</div>
			<div class="modal-body small" id="track-body">
					

			</div>
			<div class="modal-footer">
					<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
				  </div>
		  </div>
		</div>
	  </div>
	<!-- end track order modal -->

<!-- Update address modal -->
<div class="modal fade" id="updateAddressModal" tabindex="-1" role="dialog" aria-labelledby="headUpdateAddress" aria-hidden="true">
		<div class="modal-dialog" role="document">
		  <div class="modal-content">
			<div class="modal-header">
			  <h5 class="modal-title" id="headUpdateAddress">View / Update Address</h5>
			  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
				<span aria-hidden="true">&times;</span>
			  </button>
			</div>
			<form id="frm-update-address">
			<div class="modal-body modal-scroll-long" id="update-address-body">	
				<% var_update_order_address = "yes" %>	
				<!--#include virtual="/accounts/inc-cim-address-form.asp" -->
				<div id="message-address-modal"></div>
			</div>
			<div class="modal-footer">
			  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			  <button type="button" class="btn btn-purple modal-submit" id="confirm-update-address" data-id="" data-toggle="modal" data-target="#updateAddress">Update</button>
			</div>
		</form>
		  </div>
		</div>
	  </div>
	<!-- end update address modal -->

	<!-- Add-on items modal -->
<div class="modal fade" id="AddonModal" tabindex="-1" role="dialog" aria-labelledby="headAddons" aria-hidden="true">
	<div class="modal-dialog" role="document">
		<div class="modal-content">
		<div class="modal-header">
			<h5 class="modal-title" id="headAddons">Add item(s) to your order</h5>
			<button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
			</button>
		</div>
			<div class="modal-body">
By clicking the start button below, you'll be taken to our product search where you can add products to this order. You'll be able to checkout just like normal. You can cancel out of adding items at any time at the top of any page.
			</div>
		<div class="modal-footer">
			<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			<button type="button" class="btn btn-success modal-submit" id="confirm-start-addons" data-invoice="">Start adding item(s)</button>
		</div>
		</div>
	</div>
	</div>
<!-- end Add-on items modal -->

<!-- Cancel order modal -->
<div class="modal fade" id="cancelOrderModal" tabindex="-1" role="dialog" aria-labelledby="headCancelOrder" aria-hidden="true">
		<div class="modal-dialog" role="document">
		  <div class="modal-content">
			<div class="modal-header">
			  <h5 class="modal-title" id="headCancelOrder">Cancel Order</h5>
			  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
				<span aria-hidden="true">&times;</span>
			  </button>
			</div>
			<form id="frm-cancel-order">
				<div class="modal-body">
					<div id="loader-cancel-modal"></div>
					<div id="message-cancel-modal"></div>
				</div>
		</form>
			<div class="modal-footer">
			  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			  <button type="button" class="btn btn-danger modal-submit" id="confirm-cancel" data-invoice="">Cancel Order</button>
			</div>
		  </div>
		</div>
	  </div>
	<!-- end cancel order modal -->

<!-- Submit survey modal -->
<div class="modal fade" id="submitSurveyModal" tabindex="-1" role="dialog" aria-labelledby="headSubmitSurvey" aria-hidden="true">
		<div class="modal-dialog" role="document">
		  <div class="modal-content">
			<div class="modal-header">
			  <h5 class="modal-title" id="headSubmitSurvey">Submit A Survey</h5>
			  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
				<span aria-hidden="true">&times;</span>
			  </button>
			</div>
			<div class="modal-body modal-scroll-long" id="survey-body">
					<form name="frm-survey" id="frm-survey">
							<a id="survey-top"></a>
							<h6 class="m-0">How do you feel about our overall jewelry selection?<span class="text-danger"> *</span></h6>
							<div class="container p-0">
							<div class="form-check form-check-inline starrating flex-row-reverse">
									
									<input class="form-check-input star5" type="radio" name="selection" id="selection-star5" value="5" required/><label class="m-0 p-0" for="selection-star5" title="5 star" data-group="selection" data-value="5"></label>
									<input class="form-check-input star4" type="radio" name="selection" id="selection-star4"  value="4" required/><label class="m-0 p-0" for="selection-star4" title="4 star" data-group="selection" data-value="4"></label>
									<input class="form-check-input star3" type="radio" name="selection" id="selection-star3"  value="3" required/><label class="m-0 p-0" for="selection-star3" title="3 star" data-group="selection" data-value="3"></label>
									<input class="form-check-input star2" type="radio" name="selection" id="selection-star2"  value="2" required/><label class="m-0 p-0" for="selection-star2" title="2 star" data-group="selection" data-value="2" ></label>
									<input class="form-check-input star1" type="radio" name="selection" id="selection-star1"  value="1" required/><label class="m-0 p-0" for="selection-star1" title="1 star" data-group="selection" data-value="1"></label>
									<div class="invalid-feedback">
											Rating is required
									</div>
								</div>
							</div>
								<div class="textarea-selection wrapper-textarea" style="display:none">
									<label for="selectionElaborate">What jewelry would you like to see more of?<span class="text-danger font-weight-bold"> *</span></label>
									<textarea class="form-control" name="selectionElaborate" id="selectionElaborate"></textarea>
									<div class="invalid-feedback">
											Field is required
									</div>
								</div>

								<h6 class="mt-3 mb-0">How do you feel about the prices on our products?<span class="text-danger"> *</span></h6>
								<div class="container p-0">
								<div class="form-check form-check-inline starrating flex-row-reverse">
										
										<input class="form-check-input star5" type="radio" name="pricing" id="pricing-star5" value="5" required/><label class="m-0 p-0" for="pricing-star5" title="5 star" data-group="pricing" data-value="5"></label>
										<input class="form-check-input star4" type="radio" name="pricing" id="pricing-star4"  value="4" required/><label class="m-0 p-0" for="pricing-star4" title="4 star" data-group="pricing" data-value="4"></label>
										<input class="form-check-input star3" type="radio" name="pricing" id="pricing-star3"  value="3" required/><label class="m-0 p-0" for="pricing-star3" title="3 star" data-group="pricing" data-value="3"></label>
										<input class="form-check-input star2" type="radio" name="pricing" id="pricing-star2"  value="2" required/><label class="m-0 p-0" for="pricing-star2" title="2 star" data-group="pricing" data-value="2" ></label>
										<input class="form-check-input star1" type="radio" name="pricing" id="pricing-star1"  value="1" required/><label class="m-0 p-0" for="pricing-star1" title="1 star" data-group="pricing" data-value="1"></label>
										<div class="invalid-feedback">
												Rating is required
										</div>
									</div>
								</div>
									<div class="textarea-pricing wrapper-textarea" style="display:none">
										<label for="pricingElaborate">Is there another site with lower prices on the items you wanted to purchase?<span class="text-danger font-weight-bold"> *</span></label>
										<textarea class="form-control" name="pricingElaborate" id="pricingElaborate"></textarea>
										<div class="invalid-feedback">
												Field is required
										</div>
									</div>

									<h6 class="mt-3 mb-0">How would you rate the website shopping &amp; checkout experience at BAF?<span class="text-danger"> *</span></h6>
									<div class="container p-0">
									<div class="form-check form-check-inline starrating flex-row-reverse">
											
											<input class="form-check-input star5" type="radio" name="experience" id="experience-star5" value="5" required/><label class="m-0 p-0" for="experience-star5" title="5 star" data-group="experience" data-value="5"></label>
											<input class="form-check-input star4" type="radio" name="experience" id="experience-star4"  value="4" required/><label class="m-0 p-0" for="experience-star4" title="4 star" data-group="experience" data-value="4"></label>
											<input class="form-check-input star3" type="radio" name="experience" id="experience-star3"  value="3" required/><label class="m-0 p-0" for="experience-star3" title="3 star" data-group="experience" data-value="3"></label>
											<input class="form-check-input star2" type="radio" name="experience" id="experience-star2"  value="2" required/><label class="m-0 p-0" for="experience-star2" title="2 star" data-group="experience" data-value="2" ></label>
											<input class="form-check-input star1" type="radio" name="experience" id="experience-star1"  value="1" required/><label class="m-0 p-0" for="experience-star1" title="1 star" data-group="experience" data-value="1"></label>
											<div class="invalid-feedback">
													Rating is required
											</div>
										</div>
									</div>
										<div class="textarea-experience wrapper-textarea" style="display:none">
											<label for="experienceElaborate">What changes would make the website easier to use? What problems did you have?<span class="text-danger font-weight-bold"> *</span></label>
											<textarea class="form-control" name="experienceElaborate" id="experienceElaborate"></textarea>
											<div class="invalid-feedback">
													Field is required
											</div>
										</div>

								
										<h6 class="mt-3 mb-0">Did your order arrive well packaged and safe?<span class="text-danger"> *</span></h6>
										<div class="container p-0">
										<div class="form-check form-check-inline starrating flex-row-reverse">
												
												<input class="form-check-input star5" type="radio" name="packaging" id="packaging-star5" value="5" required/><label class="m-0 p-0" for="packaging-star5" title="5 star" data-group="packaging" data-value="5"></label>
												<input class="form-check-input star4" type="radio" name="packaging" id="packaging-star4"  value="4" required/><label class="m-0 p-0" for="packaging-star4" title="4 star" data-group="packaging" data-value="4"></label>
												<input class="form-check-input star3" type="radio" name="packaging" id="packaging-star3"  value="3" required/><label class="m-0 p-0" for="packaging-star3" title="3 star" data-group="packaging" data-value="3"></label>
												<input class="form-check-input star2" type="radio" name="packaging" id="packaging-star2"  value="2" required/><label class="m-0 p-0" for="packaging-star2" title="2 star" data-group="packaging" data-value="2" ></label>
												<input class="form-check-input star1" type="radio" name="packaging" id="packaging-star1"  value="1" required/><label class="m-0 p-0" for="packaging-star1" title="1 star" data-group="packaging" data-value="1"></label>
												<div class="invalid-feedback">
														Rating is required
												</div>
											</div>
										</div>
											<div class="textarea-packaging wrapper-textarea" style="display:none">
												<label for="packagingElaborate">What was wrong with the packaging?<span class="text-danger font-weight-bold"> *</span></label>
												<textarea class="form-control" name="packagingElaborate" id="packagingElaborate"></textarea>
												<div class="invalid-feedback">
														Field is required
												</div>
											</div>										

									


								<h6 class="mt-3 mb-0">How do you feel about the speed of delivery?<span class="text-danger"> *</span></h6>
								<div class="container p-0">
								<div class="form-check form-check-inline starrating flex-row-reverse">
										
										<input class="form-check-input star5" type="radio" name="delivery" id="delivery-star5" value="5" required/><label class="m-0 p-0" for="delivery-star5" title="5 star" data-group="delivery" data-value="5"></label>
										<input class="form-check-input star4" type="radio" name="delivery" id="delivery-star4"  value="4" required/><label class="m-0 p-0" for="delivery-star4" title="4 star" data-group="delivery" data-value="4"></label>
										<input class="form-check-input star3" type="radio" name="delivery" id="delivery-star3"  value="3" required/><label class="m-0 p-0" for="delivery-star3" title="3 star" data-group="delivery" data-value="3"></label>
										<input class="form-check-input star2" type="radio" name="delivery" id="delivery-star2"  value="2" required/><label class="m-0 p-0" for="delivery-star2" title="2 star" data-group="delivery" data-value="2" ></label>
										<input class="form-check-input star1" type="radio" name="delivery" id="delivery-star1"  value="1" required/><label class="m-0 p-0" for="delivery-star1" title="1 star" data-group="delivery" data-value="1"></label>
										<div class="invalid-feedback">
												Rating is required
										</div>
									</div>
								</div>
									<div class="textarea-delivery wrapper-textarea" style="display:none">
										<label for="deliveryElaborate">How long did it take your package to arrive?<span class="text-danger font-weight-bold"> *</span></label>
										<textarea class="form-control" name="deliveryElaborate" id="deliveryElaborate"></textarea>
										<div class="invalid-feedback">
												Field is required
										</div>
									</div>								


						
									<h6 class="mt-3 mb-0">If you contacted customer service about your order, how was your overall experience with that?</h6>
									<div class="container p-0">
									<div class="form-check form-check-inline starrating flex-row-reverse">
											
											<input class="form-check-input star5" type="radio" name="customerservice" id="customerservice-star5" value="5"/><label class="m-0 p-0" for="customerservice-star5" title="5 star" data-group="customerservice" data-value="5"></label>
											<input class="form-check-input star4" type="radio" name="customerservice" id="customerservice-star4"  value="4"/><label class="m-0 p-0" for="customerservice-star4" title="4 star" data-group="customerservice" data-value="4"></label>
											<input class="form-check-input star3" type="radio" name="customerservice" id="customerservice-star3"  value="3"/><label class="m-0 p-0" for="customerservice-star3" title="3 star" data-group="customerservice" data-value="3"></label>
											<input class="form-check-input star2" type="radio" name="customerservice" id="customerservice-star2"  value="2"/><label class="m-0 p-0" for="customerservice-star2" title="2 star" data-group="customerservice" data-value="2" ></label>
											<input class="form-check-input star1" type="radio" name="customerservice" id="customerservice-star1"  value="1"/><label class="m-0 p-0" for="customerservice-star1" title="1 star" data-group="customerservice" data-value="1"></label>
										</div>
									</div>
										<div class="textarea-customerservice wrapper-textarea" style="display:none">
											<label for="customerserviceElaborate">What would have made your experience with customer service better?</label>
											<textarea class="form-control" name="customerserviceElaborate" id="customerserviceElaborate"></textarea>
										</div>									


										<h6 class="mt-3 mb-0">How do you feel about your overall experience regarding this order at BAF?<span class="text-danger"> *</span></h6>
										<div class="container p-0">
										<div class="form-check form-check-inline starrating flex-row-reverse">
												
												<input class="form-check-input star5" type="radio" name="overall" id="overall-star5" value="5" required/><label class="m-0 p-0" for="overall-star5" title="5 star" data-group="overall" data-value="5"></label>
												<input class="form-check-input star4" type="radio" name="overall" id="overall-star4"  value="4" required/><label class="m-0 p-0" for="overall-star4" title="4 star" data-group="overall" data-value="4"></label>
												<input class="form-check-input star3" type="radio" name="overall" id="overall-star3"  value="3" required/><label class="m-0 p-0" for="overall-star3" title="3 star" data-group="overall" data-value="3"></label>
												<input class="form-check-input star2" type="radio" name="overall" id="overall-star2"  value="2" required/><label class="m-0 p-0" for="overall-star2" title="2 star" data-group="overall" data-value="2" ></label>
												<input class="form-check-input star1" type="radio" name="overall" id="overall-star1"  value="1" required/><label class="m-0 p-0" for="overall-star1" title="1 star" data-group="overall" data-value="1"></label>
												<div class="invalid-feedback">
														Rating is required
												</div>
											</div>
										</div>
											<div class="textarea-overall wrapper-textarea" style="display:none">
												<label for="overallElaborate">What would have made your experience better?<span class="text-danger font-weight-bold"> *</span></label>
												<textarea class="form-control" name="overallElaborate" id="overallElaborate"></textarea>
												<div class="invalid-feedback">
														Field is required
												</div>
											</div>																	


											<h6 class="mt-3 mb-0">Were all of the items correct on your order? (Quantity, matching pairs, etc)<span class="text-danger"> *</span></h6>
											<div class="container p-0">
											<div class="form-check form-check-inline">
													<div class="custom-control custom-radio custom-control-inline">
															<input type="radio" id="items-yes" name="items" class="custom-control-input" value="5" required>
															<label class="custom-control-label" for="items-yes" data-group="items"  data-value="5">Yes</label>
														</div>
														<div class="custom-control custom-radio custom-control-inline">
															<input type="radio" id="items-no" name="items" class="custom-control-input" value="0" required>
															<label class="custom-control-label" for="items-no" data-group="items"  data-value="0">No</label>
														</div>
													<div class="invalid-feedback">
															Rating is required
													</div>
												</div>
											</div>
												<div class="textarea-items wrapper-textarea" style="display:none">
													<label for="itemsElaborate">What was wrong with the item(s)?<span class="text-danger font-weight-bold"> *</span></label>
													<textarea class="form-control" name="itemsElaborate" id="itemsElaborate"></textarea>
													<div class="invalid-feedback">
															Field is required
													</div>
												</div>											

							
												<h6 class="mt-3 mb-0">Did the quality of the items you received meet your expectations?<span class="text-danger"> *</span></h6>
												<div class="container p-0">
												<div class="form-check form-check-inline">
														<div class="custom-control custom-radio custom-control-inline">
																<input type="radio" id="quality-yes" name="quality" class="custom-control-input" value="5" required>
																<label class="custom-control-label" for="quality-yes" data-group="quality"  data-value="5">Yes</label>
															</div>
															<div class="custom-control custom-radio custom-control-inline">
																<input type="radio" id="quality-no" name="quality" class="custom-control-input" value="0" required>
																<label class="custom-control-label" for="quality-no" data-group="quality"  data-value="0">No</label>
															</div>
														<div class="invalid-feedback">
																Rating is required
														</div>
													</div>
												</div>
													<div class="textarea-quality wrapper-textarea" style="display:none">
														<label for="qualityElaborate">Why were you dissatisfied with the quality?<span class="text-danger font-weight-bold"> *</span></label>
														<textarea class="form-control" name="qualityElaborate" id="qualityElaborate"></textarea>
														<div class="invalid-feedback">
																Field is required
														</div>
													</div>								
						
											

											<h6 class="mt-3 mb-0">Were all the items available that you wanted to order?<span class="text-danger"> *</span></h6>
											<div class="container p-0">
											<div class="form-check form-check-inline">
													<div class="custom-control custom-radio custom-control-inline">
															<input type="radio" id="stocklevels-yes" name="stocklevels" class="custom-control-input" value="5" required>
															<label class="custom-control-label" for="stocklevels-yes" data-group="stocklevels"  data-value="5">Yes</label>
														</div>
														<div class="custom-control custom-radio custom-control-inline">
															<input type="radio" id="stocklevels-no" name="stocklevels" class="custom-control-input" value="0" required>
															<label class="custom-control-label" for="stocklevels-no" data-group="stocklevels"  data-value="0">No</label>
														</div>
													<div class="invalid-feedback">
															Rating is required
													</div>
												</div>
											</div>
												<div class="textarea-stocklevels wrapper-textarea" style="display:none">
													<label for="stocklevelsElaborate">What items were out of stock?<span class="text-danger font-weight-bold"> *</span></label>
													<textarea class="form-control" name="stocklevelsElaborate" id="stocklevelsElaborate"></textarea>
													<div class="invalid-feedback">
															Field is required
													</div>
												</div>											
							
							<h6 class="mt-3">What new types of jewelry would you most like to see added to BAF?</h6>
								<textarea class="form-control" name="new-jewelry"></textarea>
							
							
							<h6 class="mt-3">Additional comments:</h6>
								<textarea class="form-control" name="comments"></textarea>
							
								<input type="hidden" name="invoiceid" id="surveyId" value="">						
							</form>

							<div id="message-survey-modal" class="modal-message"></div>
			</div><!-- modal scrolling div -->
			
			<div class="alert alert-warning rounded-0 p-1 m-0 small text-center d-lg-none">Scroll down to answer all questions</div>
			<div class="modal-footer">
			  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			  <button type="button" class="btn btn-purple modal-submit" id="confirm-submit-survey">Submit</button>
			</div>
		  </div>
		</div>
	  </div>
	<!-- end submit survey modal -->
	<input id="copy-photo-email" type="hidden" value="<%= rsGetUser.Fields.Item("email").Value %>">
	<input id="copy-photo-name" type="hidden" value="<%= rsGetUser.Fields.Item("customer_first").Value %>" >
	<input id="copy-photo-custid" type="hidden" value="<%= CustID_Cookie %>">
<!-- Submit a photo  modal -->
<div class="modal fade" id="submitPhotoModal" tabindex="-1" role="dialog" aria-labelledby="headSubmitPhoto" aria-hidden="true">
		<div class="modal-dialog" role="document">
		  <div class="modal-content">
			<div class="modal-header">
			  <h5 class="modal-title" id="headSubmitPhoto">Submit A Photo</h5>
			  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
				<span aria-hidden="true">&times;</span>
			  </button>
			</div>
			<div class="modal-body">
					<h6 class="title"></h6>
				Earn 3 points for each accepted photo!
				<form name="frm-submit-photo" id="frm-submit-photo" method="post" enctype="multipart/form-data" data-orderitemid="<%= var_order_detailid %>">
					<div class="alert alert-warning p-1 small">	  
					<ul>
								<li>You must own the photo and must be wearing the jewelry. No nudity.</li>
								<li>Image size no greater than 1.5MB.</li>
								<li>Must be clear, close-up, and easy to see. Low quality photos will not be approved.</li>
						  </ul>
						  By submitting your photo, you are authorizing Bodyartforms to use your photo publicly on the Bodyartforms website, social media, or for any other Bodyartforms advertising purposes. Your personal information will be kept confidential.
						</div>
						
							<input class="form-control" name="photo-filename" id="photo-filename" type="file" accept="image/jpg, image/jpeg" required>

							<input name="photo-detailid" id="photo-detailid" type="hidden" value="">
							<input name="photo-email" id="photo-email" type="hidden" value="">
							<input name="photo-name" id="photo-name" type="hidden" value="" >
							<input name="productid" id="photo-id" type="hidden" value="">
							<input name="custid" id="photo-custid" type="hidden" value="<%= CustID_Cookie %>">
							<input name="order-detailid" id="photo-orderdetailid" type="hidden" value=""> 
					</form>
				<div id="message-photo-modal"></div>
				
			</div>
			<div class="modal-footer">
			  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			  <button type="submit" class="btn btn-purple modal-submit" id="confirm-submit-photo" data-id="" data-type="">Submit</button>
			</div>
		  </div>
		</div>
	  </div>
	<!-- end submit photo modal -->


<!-- Rate / Review  modal -->
<div class="modal fade" id="reviewModal" tabindex="-1" role="dialog" aria-labelledby="headReview" aria-hidden="true">
	<div class="modal-dialog" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title" id="headReview">Rate &amp; Review</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
		<form class="needs-validation" id="frm-write-review" data-orderitemid="" novalidate>
		<div class="modal-body">
			<div id="review-body">
			<h6 class="title"></h6>
						Earn 1 point for each accepted review!
						<div class="container">
								
							<div class="form-check form-check-inline starrating d-flex justify-content-center flex-row-reverse">
									
								<input class="form-check-input star5" type="radio" name="rating" value="5" id="review-5" required/><label class="large-stars" for="review-5" title="5 star"></label>
								<input class="form-check-input star4" type="radio" name="rating" value="4" id="review-4" required/><label  class="large-stars" for="review-4" title="4 star"></label>
								<input class="form-check-input star3" type="radio" name="rating" value="3" id="review-3" required/><label  class="large-stars" for="review-3" title="3 star"></label>
								<input class="form-check-input star2" type="radio" name="rating" value="2" id="review-2" required/><label  class="large-stars" for="review-2" title="2 star" ></label>
								<input class="form-check-input star1" type="radio" name="rating" value="1" id="review-1" required/><label  class="large-stars" for="review-1" title="1 star"></label>
								<div class="invalid-feedback">
										Star rating is required
								</div>
							</div>
						
					  </div>	
					<div class="form-group">
					<textarea class="form-control" name="review" id="text-review" minlength="10" maxlength="2000" rows="7" placeholder="Type your review here" required></textarea>
					<div class="invalid-feedback">
						Text review is required
				</div>
				</div>
				<input type="hidden" value="" name="order_detail_id" id="review_order_detail_id">
			</div>
			<div id="message-review-modal" class="modal-message"></div>
		</div>
		<div class="modal-footer">
		  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
		  <button type="submit" class="btn btn-purple modal-submit" id="confirm-submit-review" data-id="" data-type="">Submit</button>
		</div>
	</form>
	  </div>
	</div>
  </div>
<!-- end Rate / Review modal -->

<!-- Report a problem  modal -->
<div class="modal fade" id="ReportProblemModal" tabindex="-1" role="dialog" aria-labelledby="headReportProblem" aria-hidden="true">
		<div class="modal-dialog" role="document">
		  <div class="modal-content">
			<div class="modal-header">
			  <h5 class="modal-title" id="headReportProblem">Report A Problem</h5>
			  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
				<span aria-hidden="true">&times;</span>
			  </button>
			</div>
			<form id="form-report-problem">
			<div class="modal-body"> 
				<div id="loader-report-problem"></div>
				<div id="message-problem-modal"></div>
			</div>
		</form>
			<div class="modal-footer">
			  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			  <button type="button" class="btn btn-purple modal-submit" id="confirm-report-problem" data-id="" data-type="">Submit</button>
			</div>
		  </div>
		</div>
	  </div>
	<!-- end report order problem modal -->


<!-- More item information modal -->
<div class="modal fade" id="MoreInfoModal" tabindex="-1" role="dialog" aria-labelledby="headMoreInfoModal" aria-hidden="true">
		<div class="modal-dialog" role="document">
		  <div class="modal-content">
			<div class="modal-header">
			  <h5 class="modal-title" id="headMoreInfoModal">Item Information</h5>
			  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
				<span aria-hidden="true">&times;</span>
			  </button>
			</div>
				<div class="modal-body">
					<div id="loader-more-information"></div>
				</div>
			<div class="modal-footer">
			  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		  </div>
		</div>
	  </div>
	<!-- end More item information modal -->

<!-- Legend for reviews and photos -->
<div class="modal fade" id="LegendModal" tabindex="-1" role="dialog" aria-labelledby="headLegendModal" aria-hidden="true">
	<div class="modal-dialog" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title" id="headLegendModal">Submitting Reviews &amp; Photos</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
			<div class="modal-body">
					
				To submit a review click on the <button class="btn btn-light btn-sm" type="button"><i class="fa fa-star-half-o fa-lg"></i></button> on any of your ordered items
				<div class="mt-2">
				To submit a photo click on the <button class="btn btn-light btn-sm" type="button"><i class="fa fa-camera fa-lg"></i></button> on any of your ordered items</div>
				<h6>Legend:</h6>
				<button class="btn btn-light btn-sm my-1" type="button">&nbsp;&nbsp;&nbsp;&nbsp;</button> No icon = Approved &amp; finished<br/>
				<button class="btn alert-success btn-sm my-1" type="button"><i class="fa fa-star-half-o fa-lg"></i></button> = Received review, pending our approval<br/>
				<button class="btn alert-success btn-sm my-1" type="button"><i class="fa fa-camera fa-lg"></i></button> = Received photo, pending our approval<br/>
				<button class="btn btn-danger btn-sm my-1" type="button"><i class="fa fa-star-half-o fa-lg"></i></button> = Declined review (You can resubmit)<br/>
				<button class="btn btn-danger btn-sm my-1" type="button"><i class="fa fa-camera fa-lg"></i></button> = Declined photo (You can resubmit)<br/>
			</div>
		<div class="modal-footer">
		  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
		</div>
	  </div>
	</div>
  </div>
<!-- end Legend for reviews and photos -->

<!-- START return policy modal -->
<div class="modal fade" id="ReturnPolicyModal" tabindex="-1" role="dialog" aria-labelledby="headReturnPolicyModal" aria-hidden="true">
	<div class="modal-dialog modal-lg" role="document">
	  <div class="modal-content">
		<div class="modal-header">
		  <h5 class="modal-title" id="headReturnPolicyModal">Return Policy</h5>
		  <button type="button" class="close" data-dismiss="modal" aria-label="Close">
			<span aria-hidden="true">&times;</span>
		  </button>
		</div>
			<div class="modal-body">
				<!--#include virtual="/misc_pages/inc-return-policy.asp"-->
			</div>
		<div class="modal-footer">
		  <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
		</div>
	  </div>
	</div>
  </div>
<!-- end return policy modal -->


<% end if   'rsGetUser.EOF %>

<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript" src="/js-pages/order-history.min.js?v=081020"></script> 
