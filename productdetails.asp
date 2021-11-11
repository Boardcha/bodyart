<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<!--#include virtual="/Connections/sql_connection.asp" -->
<%
' resetting some cart variables
session("credit_now") = 0
session("temp_shipping") = 0

ProductID = Request.QueryString("ProductID")

If IsNumeric(ProductID) AND Request.QueryString("ProductID") <> "" Then
	' Do nothing
Else
	'Exit script
	ProductID = 0
End If



set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM jewelry WHERE ProductID = ? AND ProductID <> 2424"
objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,ProductID))
Set rsProduct = objCmd.Execute()



' Get all product stats from flat products static table
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM FlatProducts WHERE ProductID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,ProductID))
Set rsProductStats = objCmd.Execute()

' Retrieve how many people have this product in their cart (and not on save for later status)
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT COUNT(cart_id) as 'currently_in_all_carts' FROM tbl_carts INNER JOIN ProductDetails ON tbl_carts.cart_detailId = ProductDetails.ProductDetailID WHERE cart_save_for_later = 0 AND cart_LastViewed > (GETDATE()- 30) AND  ProductID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,ProductID))
Set rsHowManyInCarts = objCmd.Execute()


' ----- BEGIN recently viewed items -----------
' Write to the recently viewed cookie to generate list
if IsNumeric(request.querystring("productid")) then
'response.cookies("recently_viewed") = ""
'Response.Cookies("recently_viewed").Expires = DateAdd("d",-1,now())
recents_cleaned = request.cookies("recently_viewed")
recents_cleaned = replace(recents_cleaned, "(", "")
recents_cleaned = replace(recents_cleaned, ")", "")
recents_cleaned = replace(recents_cleaned, "[", "")
recents_cleaned = replace(recents_cleaned, "]", "")
recents_cleaned = replace(recents_cleaned, """", "")

recents_array = split(recents_cleaned,",")
recents_count = uBound(recents_array) + 1

if instr(recents_cleaned, request.querystring("productid")) = 0 then ' only add if no duplicate product # is found
	' Insert first comma or not
	if recents_cleaned = "" then	
		include_comma = ""
	else
		include_comma = ","
		recents_array = split(recents_cleaned,",")
		recents_count = uBound(recents_array) + 1
		
	end if
	response.cookies("recently_viewed") = request.querystring("productid") & "" & include_comma & "" & recents_cleaned
	
	' If recents has more than 10 items then trim off the oldest #
	if recents_count >= 10 then
		trim_number = InStrRev(recents_cleaned,",") - 1
		response.cookies("recently_viewed") = Left(recents_cleaned,trim_number)
	end if
end if

'response.cookies("recently_viewed") = ""

sql_recents = ""
sql_recents_orderby = ""
i_order = 1
For Each item In recents_array
		sql_recents = sql_recents & " OR ProductID = ?"
		sql_recents_orderby = sql_recents_orderby & " WHEN ProductID = ? then " & i_order
		i_order = i_order + 1
Next

if sql_recents_orderby <> "" then
	sql_recents_orderby = " ORDER BY CASE " & sql_recents_orderby & " END"
end if


set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT ProductID, title, picture FROM jewelry WHERE ProductID = 0" & sql_recents & " " & sql_recents_orderby
	For Each item In recents_array	
			objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,item))
	Next
	For Each item In recents_array
			objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,item))
	Next
Set rsRecentlyViewed = objCmd.Execute()
end if ' if productid is numeric
' ----- END recently viewed items -----------





' Set product stats variables
if not rsProductStats.eof then
	min_gauge = rsProductStats.Fields.Item("min_gauge").Value
	max_gauge = rsProductStats.Fields.Item("max_gauge").Value
		if rsProductStats.Fields.Item("avg_rating").Value <> "" then
			avg_rating = FormatNumber(rsProductStats.Fields.Item("avg_rating").Value,1)
			var_avg_percentage = avg_rating * 20
		end if
	total_ratings = rsProductStats.Fields.Item("total_ratings").Value
	total_reviews = rsProductStats.Fields.Item("total_reviews").Value
	total_photos = rsProductStats.Fields.Item("total_photos").Value
	brand_logo = rsProductStats.Fields.Item("ProductLogo").Value
	show_brand = rsProductStats.Fields.Item("ShowTextLogo").Value
	preorder_timeframes = rsProductStats.Fields.Item("preorder_timeframes").Value
end if 

var_show_star_ratings = ""

if min_gauge <> "n/a" and min_gauge <> "" and max_gauge <> "n/a" and max_gauge <> "" then
	if min_gauge <> max_gauge then
		meta_gauge_range = Server.HTMLEncode(min_gauge & " thru " & max_gauge)
	else
		meta_gauge_range = Server.HTMLEncode(min_gauge)
	end if
end if

meta_title = "No product found"
meta_description = "No product found"
if not rsProduct.eof then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "" & _
		"SELECT TOP 50 ORD.ProductID AS ProductID, ORD2.ProductID AS bought_with, count(*) as times_bought_together, JEW.title, JEW.picture " & _
		"FROM TBL_OrderSummary AS ORD INNER JOIN TBL_OrderSummary AS ORD2 ON ORD.InvoiceID = ORD2.InvoiceID " & _
		"AND ORD.ProductID != ORD2.ProductID AND ORD.ProductID = " & ProductID & " AND ORD2.ProductID in (SELECT ProductID FROM ProductDetails WHERE " & _
		" ProductDetails.free = 0 " & _
		" AND ProductDetails.active = 1 " & _
		" AND ORD2.item_price > 0 " & _
		" AND ORD2.ProductID <> 1464 " & _
		" AND ORD2.ProductID <> 1430 " & _
		" AND ORD2.ProductID <> 1430 " & _
		" AND ORD2.ProductID <> 530 " & _
		" AND ORD2.ProductID <> 3928 " & _
		" AND ORD2.ProductID <> 15385 " & _
		" AND ORD2.ProductID <> 3611 " & _
		" AND ORD2.ProductID <> 3587 " & _
		" AND ORD2.ProductID <> 3086 " & _
		" AND ORD2.ProductID <> 3704 " & _
		" AND ORD2.ProductID <> 1649 " & _
		" AND ORD2.ProductID <> 4287 " & _
		" AND ORD2.ProductID <> 1483 " & _
		" AND ORD2.ProductID <> 3926 " & _
		" AND ORD2.ProductID <> 3803 " & _
		" AND ORD2.ProductID <> 1851 " & _
		" AND ORD2.ProductID <> 2890 " & _
		" AND ORD2.ProductID <> 2991 " & _
		"GROUP BY ProductID HAVING SUM(qty) > 0) " & _
		"LEFT JOIN Jewelry JEW ON JEW.ProductID = ORD2.ProductID " & _
		"GROUP BY ORD.ProductID, ORD2.ProductID, JEW.title, JEW.picture, JEW.active " & _
		"HAVING count(*) > 5 AND JEW.active = 1 " & _
		"ORDER BY times_bought_together DESC"
	Set rsCrossSellingItems = objCmd.Execute()

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM tbl_color_charts"
	Set rsColorCharts = objCmd.Execute()

	var_thumbs_charts = ""
	while not rsColorCharts.eof
		if instr(rsProduct.Fields.Item("ColorChart").Value, rsColorCharts.Fields.Item("chart_filename").Value) then

		var_thumbs_charts = var_thumbs_charts & "<img class='img-fluid lazyload' style='max-width: 100px;max-height: 100px' src='/images/image-placeholder.png' data-src='https://bodyartforms-products.bodyartforms.com/" & rsColorCharts.Fields.Item("chart_filename").Value & "' alt='Color Chart' title='" & rsColorCharts.Fields.Item("chart_title").Value & "' data-imgname='" &  rsColorCharts.Fields.Item("chart_filename").Value & "' data-color-chart='yes' />"
		end if

	rsColorCharts.movenext()
	wend
	rsColorCharts.requery()

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT DISTINCT i.img_id, i.product_id, i.img_full, i.img_thumb, i.img_description, i.img_sort, i.is_video, CASE WHEN p.active IS NULL THEN '1' WHEN p.active = 0 THEN '0' ELSE '1' END AS active, CASE WHEN p.img_id IS NULL THEN '0' ELSE '1' END AS detail_img_id, g.total_qty_of_variants_this_image_assigned_to " & _
	"FROM tbl_images AS i LEFT OUTER JOIN ProductDetails AS p ON i.img_id = p.img_id " & _
	"LEFT OUTER JOIN (SELECT img_id, SUM(qty) As total_qty_of_variants_this_image_assigned_to FROM ProductDetails GROUP BY img_id) g ON i.img_id=g.img_id " & _
	"WHERE i.product_id = ? ORDER BY i.img_sort"
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,ProductID))
	Set rs_getImages = objCmd.Execute()

	if rsProduct.Fields.Item("seo_meta_title").Value <> "" then
		meta_title = rsProduct.Fields.Item("seo_meta_title").Value
	else
		meta_title = rsProduct.Fields.Item("title").Value
	end if

	if rsProduct.Fields.Item("seo_meta_description").Value <> "" then
		meta_description = rsProduct.Fields.Item("seo_meta_description").Value
	else
		meta_description = "Gauges/sizes: " & meta_gauge_range
	end if

	' Only show star ratings if more than 5 people have rated the item
	if avg_rating <> "" then
		
		var_show_star_ratings = "yes"
		' Get counts for each star rating
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT TBLReviews.review_rating, COUNT(TBLReviews.review_rating) AS counts FROM TBLReviews INNER JOIN ProductDetails ON TBLReviews.DetailID = ProductDetails.ProductDetailID GROUP BY ProductDetails.ProductID, TBLReviews.review_rating HAVING (ProductDetails.ProductID = ?) AND (TBLReviews.review_rating <> 0) ORDER BY review_rating DESC"
		objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,ProductID))
		set rs_StarCounts = objCmd.Execute()
	end if	' avg_rating <> ""

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT *, ISNULL(Gauge, '') + ' ' + ISNULL(Length, '') + ' ' + ISNULL(ProductDetail1, '') as OptionTitle FROM ProductDetails WHERE ProductID = ? AND active = 1 ORDER BY item_order ASC, Price ASC"
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,ProductID))
	Set rs_getDropDownItems = objCmd.Execute()
	total_drop_down_items = rs_getDropDownItems.RecordCount
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT *, ISNULL(Gauge, '') + ' ' + ISNULL(Length, '') + ' ' + ISNULL(ProductDetail1, '') as OptionTitle  FROM ProductDetails WHERE ProductID = ? AND active = 1 and qty > 0 ORDER BY item_order ASC, Price ASC"
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,ProductID))

	set rsGetActiveItems = Server.CreateObject("ADODB.Recordset")
	rsGetActiveItems.CursorLocation = 3 'adUseClient
	rsGetActiveItems.Open objCmd
	var_totalActiveitems = rsGetActiveItems.RecordCount
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.Gauge FROM ProductDetails INNER JOIN TBL_GaugeOrder ON ProductDetails.Gauge = TBL_GaugeOrder.GaugeShow WHERE (ProductDetails.ProductID = ?) AND (ProductDetails.active = 1) AND (ProductDetails.qty > 0) GROUP BY Gauge, GaugeOrder ORDER BY GaugeOrder"
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,ProductID))

	set rsGaugeFilter = Server.CreateObject("ADODB.Recordset")
	rsGaugeFilter.CursorLocation = 3 'adUseClient
	rsGaugeFilter.Open objCmd
	total_gauges = rsGaugeFilter.RecordCount

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.Gauge AS gauge, COUNT(*) AS gauge_total, TBL_GaugeOrder.GaugeOrder FROM  TBL_GaugeOrder INNER JOIN ProductDetails ON TBL_GaugeOrder.GaugeShow = ProductDetails.Gauge RIGHT OUTER JOIN TBLReviews ON ProductDetails.ProductDetailID = TBLReviews.DetailID WHERE (TBLReviews.ProductID = ?) AND (TBLReviews.status = N'accepted') AND (ProductDetails.Gauge IS NOT NULL) GROUP BY ProductDetails.Gauge, TBL_GaugeOrder.GaugeOrder ORDER BY TBL_GaugeOrder.GaugeOrder"
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,12,ProductID))
	Set rs_reviews_gauge_dropdown = objCmd.Execute()

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.ProductDetail1 AS color, COUNT(*) AS color_total FROM ProductDetails RIGHT OUTER JOIN TBLReviews ON ProductDetails.ProductDetailID = TBLReviews.DetailID 	WHERE (TBLReviews.ProductID = ?) AND (TBLReviews.status = N'accepted') AND (ProductDetails.ProductDetail1 IS NOT NULL) GROUP BY ProductDetails.ProductDetail1"
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,12,ProductID))
	Set rs_reviews_color_dropdown = objCmd.Execute()

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.Gauge AS gauge, COUNT(*) AS gauge_total FROM TBL_GaugeOrder INNER JOIN ProductDetails ON TBL_GaugeOrder.GaugeShow = ProductDetails.Gauge RIGHT OUTER JOIN TBL_PhotoGallery ON ProductDetails.ProductDetailID = TBL_PhotoGallery.DetailID WHERE (ProductDetails.Gauge IS NOT NULL) AND (TBL_PhotoGallery.ProductID = ?) AND (TBL_PhotoGallery.status = 1) GROUP BY ProductDetails.Gauge, TBL_GaugeOrder.GaugeOrder ORDER BY TBL_GaugeOrder.GaugeOrder"
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,12,ProductID))
	Set rs_photos_gauge_dropdown = objCmd.Execute()

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.ProductDetail1 AS color, COUNT(*) AS color_total FROM ProductDetails RIGHT OUTER JOIN TBL_PhotoGallery ON ProductDetails.ProductDetailID = TBL_PhotoGallery.DetailID WHERE (ProductDetails.ProductDetail1 IS NOT NULL) AND (TBL_PhotoGallery.ProductID = ?) AND (TBL_PhotoGallery.status = 1) GROUP BY ProductDetails.ProductDetail1"
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,12,ProductID))
	Set rs_photos_color_dropdown = objCmd.Execute()


	' Get the top 5 newest reviews to display at bottom
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP 5 TBLReviews.ProductID, TBLReviews.ReviewID, TBLReviews.review, TBLReviews.status, TBLReviews.name, TBLReviews.date_posted, TBLReviews.comments, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.ProductDetail1 FROM TBLReviews LEFT OUTER JOIN ProductDetails ON TBLReviews.DetailID = ProductDetails.ProductDetailID WHERE TBLReviews.ProductID = ? AND TBLReviews.status = N'accepted' ORDER BY ReviewID DESC" 
	objCmd.Prepared = true
	objCmd.Parameters.Append objCmd.CreateParameter("param1", 3, 1, 10, ProductID) ' adDouble
	Set rsGetReview = objCmd.Execute



	' Get all the photos for the product to display
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP(20) TBL_PhotoGallery.PhotoID, TBL_PhotoGallery.ProductID, TBL_PhotoGallery.DetailID, TBL_PhotoGallery.filename, TBL_PhotoGallery.thumb_filename, TBL_PhotoGallery.description, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.ProductDetail1, TBL_PhotoGallery.status, jewelry.title, TBL_PhotoGallery.DateSubmitted FROM TBL_PhotoGallery INNER JOIN ProductDetails ON TBL_PhotoGallery.DetailID = ProductDetails.ProductDetailID INNER JOIN jewelry ON TBL_PhotoGallery.ProductID = jewelry.ProductID WHERE status = 1 AND TBL_PhotoGallery.ProductID = ? ORDER BY PhotoID DESC"
	objCmd.Prepared = true
	objCmd.Parameters.Append objCmd.CreateParameter("param1", 3, 1, 10, ProductID) ' adDouble
	Set rsGetPhotos = objCmd.Execute

	' Get all gauges for description area available in
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT stuff((SELECT ', ' + ISNULL(d.Gauge, N'') FROM ProductDetails AS d INNER JOIN TBL_GaugeOrder AS g ON ISNULL(d.Gauge, N'') = ISNULL(g.GaugeShow, N'') WHERE (d.ProductID = ?) AND (d.active = 1) AND (d.Gauge <> 'n/a') AND (d.Gauge <> ' ') GROUP BY ISNULL(d.Gauge, N''), g.GaugeOrder ORDER BY g.GaugeOrder FOR XML PATH('')), 1, 2, '') as 'sizes_offered'"
	'=========== TESTING VERSION TO MAKE GAUGE A LINK'
	'	objCmd.CommandText = "SELECT stuff((SELECT ', ' + ISNULL(d.Gauge, N'') FROM ProductDetails AS d INNER JOIN TBL_GaugeOrder AS g ON ISNULL(d.Gauge, N'') = ISNULL(g.GaugeShow, N'') WHERE (d.ProductID = ?) AND (d.active = 1) AND (d.Gauge <> 'n/a') AND (d.Gauge <> ' ') GROUP BY ISNULL(d.Gauge, N''), g.GaugeOrder ORDER BY g.GaugeOrder FOR XML PATH('')), 1, 2, '') as 'sizes_offered'"
	objCmd.Parameters.Append objCmd.CreateParameter("param1", 3, 1, 10, ProductID) ' adDouble
	Set rsSizesOffered = objCmd.Execute

	if NOT rsSizesOffered.eof then
		if rsSizesOffered.Fields.Item("sizes_offered").Value <> "" then
			var_sizes_offered = "<li>Sizes offered: " & rsSizesOffered.Fields.Item("sizes_offered").Value & "</li>"
		end if
		
	end if


	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM customers WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
	Set rsGetUser = objCmd.Execute()
	
	Flagged = ""
	var_customer_id = 0
	var_customer_name = ""
	var_customer_email = ""
	If Not rsGetUser.EOF Then

		var_customer_id = rsGetUser.Fields.Item("customer_ID").Value
		var_customer_name = rsGetUser.Fields.Item("customer_first").Value
		var_customer_email = rsGetUser.Fields.Item("email").Value
		
		IF rsGetUser.Fields.Item("Flagged").Value = "Y" then
		Flagged = "yes"
		end if 
		
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT  ProductDetails.ProductDetail1, wishlist.itemDetailID, wishlist.desired, wishlist.comments, wishlist.dateadded, jewelry.jewelry, wishlist.WishlistID, ProductDetails.active, jewelry.ProductID, ProductDetails.Gauge, ProductDetails.Length FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID INNER JOIN wishlist ON ProductDetails.ProductDetailID = wishlist.itemDetailID INNER JOIN customers ON wishlist.custID = customers.customer_ID WHERE jewelry.ProductID = ? AND custID = ? AND ProductDetails.active <> 0"
		objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,ProductID))
		objCmd.Parameters.Append(objCmd.CreateParameter("customerID",3,1,10,CustID_Cookie))
		Set rsGetWishlist = objCmd.Execute()
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM TBL_WishlistCategories WHERE Wishlist_CustomerID = ? ORDER BY WishlistName ASC"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
		Set rsGetCategory = objCmd.Execute()
		
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT TOP (100) PERCENT SUM(qty) AS qty, Gauge + ' ' + Length + ' ' + ProductDetail1 + ' '  + PreOrder_Desc AS Gauge FROM QRY_OrderDetails WHERE (customer_ID = ?) AND (ProductID = ?) AND (ship_code = 'paid') GROUP BY ProductDetail1, Gauge, Length, ProductDetailID, PreOrder_Desc"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
		objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,ProductID))
		Set rsGetOrderHistory = objCmd.Execute()
		
	End if  ' Not rsGetUser.EOF
	
	' Set materials variables
		var_materials = ""
	if rsProduct.Fields.Item("material").Value <> "" then
	
	material_main_array = split(rsProduct.Fields.Item("material").Value," , ")
		For Each strItem In material_main_array
			if strItem <> "" and strItem <> "null " and strItem <> " " then 

			var_materials =  trim(var_materials) & "<a href='/products.asp?material=" & Trim(strItem) & "' class='text-info'>" & Trim(strItem) & "</a>,&nbsp;"
				'response.write trim(strItem) & "<br/>"
				'response.write trim(var_materials) & "<br/>"
			end if 			
		Next
	end if ' rsProduct("material") <> ""

		if var_materials <> "" then
			var_materials = Replace(var_materials, "Metal,&nbsp;", "")
			var_materials = Replace(var_materials, "Organic,&nbsp;", "")
			
			if var_materials <> "" then
				' trim last comma
				var_materials = "<li>Material(s): " & LEFT(var_materials, (LEN(var_materials)-7)) & " <a href=""materials.asp""><i class=""fa fa-question-circle""></i></a></li>"
			end if
		end if

	' Set page title
	var_title = rsProduct.Fields.Item("title").Value
	if rsProduct.Fields.Item("type").Value = "onetime" OR rsProduct.Fields.Item("type").Value = "limited" then
		var_title = "Limited " & var_title 
	elseif rsProduct.Fields.Item("type").Value = "Discontinued" then
		var_title = "Discontinued " & var_title 
	elseif rsProduct.Fields.Item("type").Value = "Clearance" then
		var_title = "Clearance " & var_title
	end if ' Set page title

	' Set variables for item status --------------
	var_status = ""
	if rsProduct.Fields.Item("type").Value = "One time buy" then
		var_status = "ONE OF A KIND"
	elseif rsProduct.Fields.Item("type").Value = "limited" then
		var_status = "LIMITED"
	elseif rsProduct.Fields.Item("type").Value = "Discontinued" then
		var_status = "LAST CHANCE"
	end if

	' Set variable for date added -------------
	var_new_date = ""
	if rsProduct.Fields.Item("new_page_date").Value <= date()+21 AND rsProduct.Fields.Item("new_page_date").Value > date()-70 then
		
		var_new_date = MonthName(Month(rsProduct.Fields.Item("new_page_date").Value),1) & " " & Day(rsProduct.Fields.Item("new_page_date").Value)
		
	end if
	

	' Set some variables based off the detail items
		WeightNotice = ""
		ShowOutItems = ""
		var_count_details = 0
		display_add_to_cart = 0
	While NOT rs_getDropDownItems.EOF
		
		var_count_details = var_count_details + 1
		
		' Detect whether there are any out of stock items (to show or hide tab)
		if (rs_getDropDownItems.Fields.Item("qty").Value) <= "0" then 
			ShowOutItems = "yes" 
		end if
		
		if (rs_getDropDownItems.Fields.Item("qty").Value) > "0" then 
			display_add_to_cart = 1 
		end if
		
		if rs_getDropDownItems.Fields.Item("weight").Value > 4 then
			WeightNotice = "yes"
		end if
		
		if lcase(rs_getDropDownItems.Fields.Item("wearable_material").Value) = "stone" OR lcase(rs_getDropDownItems.Fields.Item("wearable_material").Value) = "wood" OR lcase(rs_getDropDownItems.Fields.Item("wearable_material").Value) = "amber" OR lcase(rs_getDropDownItems.Fields.Item("wearable_material").Value) = "bamboo" OR lcase(rs_getDropDownItems.Fields.Item("wearable_material").Value) = "bone" OR lcase(rs_getDropDownItems.Fields.Item("wearable_material").Value) = "horn" OR lcase(rs_getDropDownItems.Fields.Item("wearable_material").Value) = "jet" OR lcase(rs_getDropDownItems.Fields.Item("wearable_material").Value) = "palm seed" OR lcase(rs_getDropDownItems.Fields.Item("wearable_material").Value) = "shell" OR lcase(rs_getDropDownItems.Fields.Item("wearable_material").Value) = "vegan" OR lcase(rs_getDropDownItems.Fields.Item("wearable_material").Value) = "fossils" OR lcase(rs_getDropDownItems.Fields.Item("wearable_material").Value) = "fossilized bone" then
			size_variation_notice = "<span class=""d-block mt-1"">This item is crafted from organic material and the <strong>gauge can vary +/- up to 0.5mm</strong></span>"
		 end if
			
	rs_getDropDownItems.MoveNext()
	Wend
	rs_getDropDownItems.Requery()

	' Set all misc or simple variables
	' Set flares variables
		var_flares = ""
	if rsProduct.Fields.Item("flare_type").Value <> "" then
	
	material_main_array = split(rsProduct.Fields.Item("flare_type").Value," , ")
		For Each strItem In material_main_array
			if strItem <> "" and strItem <> "null " and strItem <> " " then 

			var_flares = trim(var_flares) + Trim(strItem) & ",&nbsp;"
			'	response.write trim(strItem) & "<br/>"
			'	response.write trim(var_flares) & "<br/>"
			end if 			
		Next
	end if ' rsProduct("flare_type") <> ""

		if var_flares <> "" then		
			if var_flares <> "" then
				' trim last comma
				var_flares = "<li>Flares: " & LEFT(var_flares, (LEN(var_flares)-7)) & "</li>"
			end if
		end if
	
	var_threading = trim(replace(replace(replace(rsProduct.Fields.Item("internal").Value, ",", ""), "n/a", ""), "null", ""))
	if var_threading <> "" then
		var_threading_type = "<li>" & var_threading & "</li>"
	end if

	var_worn_in_cleaned = rsProduct.Fields.Item("piercing_type").Value
if var_worn_in_cleaned <> "" then
	var_worn_in_cleaned = trim(var_worn_in_cleaned)
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "piercing_type:" , "")

	if instr(var_worn_in_cleaned, "Helix") > 0 or instr(var_worn_in_cleaned, "Conch") > 0 or instr(var_worn_in_cleaned, "Daith") > 0 or instr(var_worn_in_cleaned, "Industrial") > 0 or instr(var_worn_in_cleaned, "Rook") > 0 or instr(var_worn_in_cleaned, "Snug") > 0 or instr(var_worn_in_cleaned, "Tragus") > 0 then
		var_wornin_replaced_terms = ",&nbsp;Ear&nbsp;cartilage"
	end if

	if instr(var_worn_in_cleaned, "Bites") > 0 or instr(var_worn_in_cleaned, "Bridge") > 0 or instr(var_worn_in_cleaned, "Cheek") > 0 or instr(var_worn_in_cleaned, "Jestrum") > 0 or instr(var_worn_in_cleaned, "Philtrum") > 0 or instr(var_worn_in_cleaned, "Vertical labret") > 0 then
		var_wornin_replaced_terms = var_wornin_replaced_terms & ",&nbsp;Face"
	end if

	if instr(var_worn_in_cleaned, "Basic ear piercing") > 0 then
		var_wornin_replaced_terms = var_wornin_replaced_terms & ",&nbsp;Ear&nbsp;lobe"
	end if

	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Ampallang" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Apadravya" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Clitoris" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Christina" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Dydoe" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Foreskin" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Fourchette" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Frenum" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Guiche" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Horizontal hood" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Labia" , "-")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Prince Albert" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Scrotum" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Vertical hood" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "None" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Microdermal" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Surface" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Tongue web" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Anti-tragus" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Basic ear piercing" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Conch" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Daith" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Helix" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Rook" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Snug" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Tragus" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Stretched lobe" , "Ear&nbsp;lobe")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Bites" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Bridge" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Cheek" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Jestrum" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Philtrum" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Vertical labret" , "")
	var_worn_in_cleaned = replace(var_worn_in_cleaned, "Navel" , "Belly&nbsp;button")
	var_worn_in_cleaned = trim(var_worn_in_cleaned) & trim(var_wornin_replaced_terms)
	
	' Removes any duplicates spaces back to back and replaces with just one
	Dim regEx
	Set regEx = New RegExp
	regEx.Global = true
	regEx.IgnoreCase = True
	regEx.Pattern = "\s{2,}"
	var_worn_in_cleaned = Trim(regEx.Replace(var_worn_in_cleaned, " "))

	var_worn_in_cleaned = replace(var_worn_in_cleaned, " " , ", ")
	' If a comma is the first character strip it out below
	if instr(var_worn_in_cleaned, ",") = 1 then
		var_worn_in_cleaned =  right(var_worn_in_cleaned,len(var_worn_in_cleaned)-1)
	end if

	wornin_main_array = split(var_worn_in_cleaned,",")
	For Each strItem In wornin_main_array
		strItem = replace(strItem, "&nbsp;" , " ")
		var_worn_in =  trim(var_worn_in) & "<a href='/products.asp?piercing=" & Trim(server.htmlencode(strItem)) & "' class='text-info'>" & Trim(strItem) & "</a>,&nbsp;"
			'response.write trim(strItem) & "<br/>"
			'response.write trim(var_worn_in_cleaned) & "<br/>"	
	Next

	if var_worn_in_cleaned <> "" then
	' Trim last 7 characters off string ,&nbsp;
		var_worn_in = "<li>Worn In: " & left(var_worn_in,len(var_worn_in)-7) & "</li>"
	end if
	
end if ' var_worn_in_cleaned <> ""
	
		var_sizing_type = ""
	if InStr(rsProduct.Fields.Item("jewelry").Value,"septum") > 0 then
		var_sizing_type = "septum"
	end if
	if InStr(rsProduct.Fields.Item("jewelry").Value,"captive") > 0 then
		var_sizing_type = "captive"
	end if
	if InStr(rsProduct.Fields.Item("jewelry").Value,"finger-ring") > 0 then
		var_sizing_type = "finger"
	end if
	
		var_product_notice = ""
		var_mini_notice = ""
	' BUILD PRODUCT NOTICES
	if rsProduct.Fields.Item("SaleExempt").Value = 1 then
		var_product_notice = var_product_notice & "<div class=""alert alert-warning mt-1""><strong>This item is exempt from any additional sales or coupons</strong> (except for the 10% preferred customer discount).</div>"
	 end if
	' Exempt from returning for cleansers or tools items
	 If WeightNotice = "yes" then	
		var_product_notice = var_product_notice & "<div class=""alert alert-warning mt-1""><strong>This item requires  priority mail shipping</strong> (USA only) or UPS shipping due to its weight, fragility, or odd shape. Free shipping is excluded from this product.</div>"
	end if ' display weight notice	
	if size_variation_notice <> "" then
		var_product_notice = var_product_notice & size_variation_notice
	 end if
	 if instr(lcase(rsProduct.Fields.Item("material").Value),"wood") > 1 then
		var_product_notice = var_product_notice & "<div class=""alert alert-warning mt-1""><strong>SPECIAL CARE:</strong> This item is made with wood. We recommend taking wood jewelry out before showering as water can ruin it.</div>"
	 end if
	 if rsProduct("brandname") = "steel and silver" AND instr(lcase(var_threading_type), "internally") AND instr(lcase(var_sizes_offered), "14g") <= 0 then
		var_product_notice = var_product_notice & "<div class=""alert alert-warning mt-1""><strong>This jewelry is not compatible with any other ends</strong><br><a href='/products.asp?keywords=silver&jewelry=curved&jewelry=labret&jewelry=barbell&gauge=16g&material=316L+Stainless+Steel&material=Titanium&price=0%3B100&threading=Internally+threaded' target='_blank'>Click here</a> for compatible ends.</div>"
	end if
	 

	
		var_pair_status = ""
	'SET WHETHER SOLD AS PAIR OR SINGLE
	if (rsProduct.Fields.Item("pair").Value) = "yes" then
		var_pair_status = "PAIR"
	else
		var_pair_status = "SINGLE"
	end if
	
		var_qty_default = "1"
	' SET DEFAULT QTY FOR INPUT BOX
	 if InStr(rsProduct.Fields.Item("jewelry").Value,"plugs") > 0 and rsProduct.Fields.Item("pair").Value <> "yes" then
			var_qty_default = "2"
	 end if
	 
	If show_brand = "Y" then 
		if brand_logo <> "" then
			var_brand_logo = "<li>Brand:</li><li style=""list-style-type: none""><a href=""products.asp?brand=" & rsProduct.Fields.Item("brandname").Value & """><img src=""images/" &  brand_logo  & """ alt=""View more products by this brand"" title=""View more products by this brand"" class=""img-fluid details-brand-logo"" /></a></li>"
		end if
	end if
	if rsProduct("country_origin") <> "" then
		origin_country = ""
	end if
		
		var_regular_stock = "yes"
	if (rsProduct.Fields.Item("type").Value) <> "limited" and (rsProduct.Fields.Item("type").Value) <> "One time buy" and (rsProduct.Fields.Item("type").Value) <> "Clearance" and (rsProduct.Fields.Item("type").Value) <> "Discontinued" then
		var_regular_stock = ""
	end if

	'=====   JSON MARKUP FOR GOOGLE TO DISPLAY REVIEW DATA
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP(20) TBLReviews.ProductID, TBLReviews.ReviewID, TBLReviews.review, TBLReviews.review_rating, TBLReviews.status, TBLReviews.name, TBLReviews.date_posted, TBLReviews.comments, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.ProductDetail1 FROM TBLReviews LEFT OUTER JOIN ProductDetails ON TBLReviews.DetailID = ProductDetails.ProductDetailID WHERE review <> '' AND review_rating > 0 AND TBLReviews.ProductID = ? AND (TBLReviews.status = N'accepted' or TBLReviews.status IS NULL) ORDER BY ReviewID DESC" 
	objCmd.Parameters.Append objCmd.CreateParameter("productid", 3, 1, 12, ProductID)

set rsJsonReviews = Server.CreateObject("ADODB.Recordset")
rsJsonReviews.CursorLocation = 3 'adUseClient
rsJsonReviews.Open objCmd
total_json_reviews = rsJsonReviews.RecordCount

end if ' not rsProduct.eof


	page_title = meta_title
	page_description = meta_description
	var_meta_productdetails = "yes"
%>

<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<% If not rsProduct.eof then %>
<script type="text/javascript">
// GA4 GTM push
window.dataLayer = window.dataLayer || [];
window.dataLayer.push({
	'event': 'view_item',
	'ecommerce': {
		'items': [{
			'item_name': '<%= rsProduct.Fields.Item("title").Value %>',
        	'item_category': '<%= trim(rsProduct.Fields.Item("jewelry").Value) %>',
        	'item_brand': '<%= rsProduct.Fields.Item("brandname").Value %>'
      }]
  }
});

// Standard UA GTM push
// GTM Product view
window.dataLayer = window.dataLayer || [];
window.dataLayer.push({
  event: 'ua_viewproduct',
  ecommerce: {
    detail: {
      products: [{
        id: '<%= rsProduct.Fields.Item("ProductID").Value %>',
        name: '<%= rsProduct.Fields.Item("title").Value %>',
        category: '<%= trim(rsProduct.Fields.Item("jewelry").Value) %>',
        brand: '<%= rsProduct.Fields.Item("brandname").Value %>'
      }]
    }
  }
});

// Klaviyo view item push
var _learnq = _learnq || [];
	var item = {
	  "ProductName": '<%= rsProduct.Fields.Item("title").Value %>',
	  "ProductID": '<%= rsProduct.Fields.Item("ProductID").Value %>',
	  "ImageURL": 'https://bodyartforms-products.bodyartforms.com/<%= rsProduct.Fields.Item("largepic").Value %>',
	  "URL": 'https://bodyartforms.com/productdetails.asp?productid=<%= rsProduct.Fields.Item("ProductID").Value %>',
	  "Brand": '<%= rsProduct.Fields.Item("brandname").Value %>'
	};
	_learnq.push(["track", "Viewed Product", item]);


// Regular Google UA Ecommerce data layer push for GTM - add to cart
window.onload=function(){
var button_addcart = document.getElementById('btn-add-cart');
		button_addcart.addEventListener("click",function(e){

			var qty = document.getElementById("add-qty").value;
			var_detailid = $('.add-cart:checked').val();
			var actual_price = $('.add-cart:checked').attr('data-actual-price');

			var variant = document.querySelector("input[name=add-cart]:checked");
			variant = variant.getAttribute("data-variant");
			if (document.cookie.indexOf('currency') > -1 ) {
				var currency = getCookie("currency");
			} else {
				var currency = 'USD'
			}
			window.dataLayer.push({
			event: 'ua_add_to_cart',
			ecommerce: {
				'currencyCode': currency,
				add: {
				products: [{
					id: '<%= rsProduct.Fields.Item("ProductID").Value %>',
					name: '<%= rsProduct.Fields.Item("title").Value %>',
					category: '<%= trim(rsProduct.Fields.Item("jewelry").Value) %>',
					variant: variant,
					brand: '<%= rsProduct.Fields.Item("brandname").Value %>',
					quantity: qty
					
				}]
				}
			}
			});

			//Klaviyo Add to cart push
			_learnq.push(["track", "Added to Cart", {
			"$value": actual_price * qty,
			"AddedItemProductName": '<%= rsProduct.Fields.Item("title").Value %>',
			"AddedItemProductID": '<%= rsProduct.Fields.Item("ProductID").Value %>',
			"AddedItemSKU": var_detailid,
			"AddedItemImageURL": 'https://bodyartforms-products.bodyartforms.com/<%= rsProduct.Fields.Item("largepic").Value %>',
			"AddedItemURL": 'https://bodyartforms.com/productdetails.asp?productid=<%= rsProduct.Fields.Item("ProductID").Value %>',
			"AddedItemPrice": actual_price,
			"AddedItemQuantity": qty,
			"CheckoutURL": "https://bodyartforms.com/checkout.asp"
		}]);	

	}); // listener for click button_addcart

} // Run after window finishes loading
</script>

<% end if %>
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<link rel="stylesheet" href="/CSS/jquery.fancybox.min.css" />
<link rel="stylesheet" type="text/css" href="/CSS/slick.css"/>

<% If rsProduct.eof or Flagged = "yes" then %>
		<div class="alert alert-danger border border-warning">Product not found</div>
	<% else %>

<div class="display-5">
	<%= var_title %>
</div>
<div>
	<% if var_show_star_ratings = "yes" then %>
	<a class="text-dark ml-1" href="#reviews">
			<span class="rating-box">
					<span class="rating" style="width:<%= var_avg_percentage %>%"></span>
				</span>
		<% if avg_rating <> "" then %>
			<span class="ml-1">(<%= avg_rating %>)</span>
		<% end if %>
	</a>
	<% end if %>
	<% if var_new_date <> "" then %>
		<span class="badge badge-primary">
			NEW <span class="font-weight-normal"><%= var_new_date %>
		</span>
	</span>
	<% end if %>
	<% if var_status <> "" then %>
		<span class="badge badge-info">
			<%= var_status %>
		</span>
	<% end if %>
	<% if rsProduct.Fields.Item("secret_sale").Value = 1 AND session("secret_sale") = "yes" then %>
	<span class="badge badge-danger">
		SECRET SALE
	</span>
	<% end if %>
<div>
	<% If rsProduct.Fields.Item("customorder").Value = "yes" then %>
		<div class="alert alert-warning mt-2">	
			<span class="font-weight-bold"><%= preorder_timeframes %> to receive</span> while it's being made to your specifications (regardless of which shipping method you choose). We will ship your order in full once we receive your custom piece.
		</div>
	<% end if %>		
<section class=" mt-4">
	<div class="container-fluid pl-0 pr-2 pb-4">
		<div class="row">
			<div class="col-12 col-lg-6 col-xl-5 mb-3">
					<div class="row row-list">
							<% if NOT rs_getImages.eof or var_thumbs_charts <> "" then
								var_show_thumbnails = "yes"
							end if	
							if rsProduct.Fields.Item("ColorChart").Value <> "" then
									var_chart_exists = "yes"
							end if
							if NOT rs_getImages.eof then 
								var_images_found = "yes"
							end if
							if var_show_thumbnails = "yes" then
							%>
							<div class="m-0 pr-0 col-2 col-md-2 col-lg-2 col-xl-2 col-break1600-1 baf-carousel-vertical col-img-thumbs" id="vert-thumb-carousel">
								<img  class="img-fluid img-thumb lazyload" src="/images/image-placeholder.png" data-src="https://bodyartforms-products.bodyartforms.com/<%=rsProduct.Fields.Item("picture").Value %>" alt="Main photo" title="Main photo"  data-imgid="thumb_img_0" data-imgname="<%= rsProduct.Fields.Item("largepic").Value %>" id="img_thumb_0" data-id="0" data-color-chart="no" data-unassigned="no" />
							<%							
							end if 


							
							' Loop through extra images if there are any
							if NOT rs_getImages.eof then 
							
								while NOT rs_getImages.eof

									'To show an additional image (thumbnail) as greyed out, all variants the image assigned must be out of stock
									If rs_getImages("total_qty_of_variants_this_image_assigned_to") <=0 Then isThumbnailQtyZero = true Else isThumbnailQtyZero = false
										
									if rs_getImages.Fields.Item("active").Value = 1 then
										
										if rs_getImages.Fields.Item("detail_img_id").Value = 1 then
											img_assigned = "yes"
										else
											img_assigned = "no"
										end if			
									
										var_img_thumb_alt = ""
										if rs_getImages.Fields.Item("img_description").Value <> "" then
											var_img_thumb_alt = rs_getImages.Fields.Item("img_description").Value
										end if
										%>	
										<%if rs_getImages.Fields.Item("is_video").Value = 1 then%>
											<div class="video-thumbnail">
												<img class="img-fluid img-thumb lazyload" src="/images/image-placeholder.png" data-src="https://bodyartforms-products.bodyartforms.com/<%=rs_getImages.Fields.Item("img_thumb").Value %>" alt="Thumbnail" title="<%= var_img_thumb_alt %>"  data-imgname="<%=(rs_getImages.Fields.Item("img_full").Value)%>" <% if img_assigned = "yes" then %> data-imgid="thumb_img_<%=rs_getImages.Fields.Item("img_id").Value %>" id="img_thumb_<%=rs_getImages.Fields.Item("img_id").Value %>" data-id="<%=rs_getImages.Fields.Item("img_id").Value %>" data-unassigned="no" <% else %> data-unassigned="yes" <% end if %> data-color-chart="no" />
												<img src="/images/play-icon.png" class="play-icon" />
											</div>	
										<%else%>	
										<div class="image-thumbnail <%If isThumbnailQtyZero Then%>diagonal-line<%End If%>">										
											<img <%If isThumbnailQtyZero Then%>style="opacity:.3"<%End If%> class="img-fluid img-thumb lazyload" src="/images/image-placeholder.png" data-src="https://bodyartforms-products.bodyartforms.com/<%=rs_getImages.Fields.Item("img_thumb").Value %>" alt="Thumbnail" title="<%= var_img_thumb_alt %>"  data-imgname="<%=(rs_getImages.Fields.Item("img_full").Value)%>" <% if img_assigned = "yes" then %> data-imgid="thumb_img_<%=rs_getImages.Fields.Item("img_id").Value %>" id="img_thumb_<%=rs_getImages.Fields.Item("img_id").Value %>" data-id="<%=rs_getImages.Fields.Item("img_id").Value %>" data-unassigned="no" <% else %> data-unassigned="yes" <% end if %> data-color-chart="no" />
										</div>
										<%end if%>
									<%
									end if ' only show if item is active
									rs_getImages.MoveNext()
								wend
								rs_getImages.ReQuery()
							end if
							%>
					
							<%= var_thumbs_charts %>
							<%
							if var_show_thumbnails = "yes" then		
							%>
							</div><!-- end product-thumbnails -->
							<% 
							end if 

							if var_show_thumbnails = "yes" then
								main_col_sizes = "col-10 col-md-9 col-lg-10 col-xl-10 col-break1600-11"
							else
								main_col_sizes = "col"
							end if
							%>
							<div class="m-0 <%= main_col_sizes %> col-img-main">
								<div class="slider-main-image baf-carousel" style="max-width:650px;max-height: 550px" >
									<a class="position-relative pointer" data-fancybox="product-images" data-caption="Main Photo" href="https://bodyartforms-products.bodyartforms.com/<%=(rsProduct.Fields.Item("largepic").Value)%>" id="img_id_0">
										<span class="position-absolute badge badge-secondary p-2 rounded-0" style="top:0;left:0;z-index:200"><i class="fa fa-search fa-lg"></i></span>
										<img class="img-fluid" height="550px" width="550px" src="https://bodyartforms-products.bodyartforms.com/<%=(rsProduct.Fields.Item("largepic").Value)%>" alt="<%= rsProduct.Fields.Item("title").Value %>" style="max-height: 550px;" />
									<div class="text-center small text-dark">Main Photo</div>
									</a>
										
								<% while NOT rs_getImages.eof
								if rs_getImages.Fields.Item("active").Value = 1 then
								var_img_title = ""
								if rs_getImages.Fields.Item("img_description").Value <> "" then
									var_img_title = rs_getImages.Fields.Item("img_description").Value
								end if
								%>
								<%if rs_getImages.Fields.Item("is_video").Value = 1 then%>
								<a class="position-relative pointer" data-fancybox="product-images" data-caption="<%= var_img_title %>" href="#video_<%=rs_getImages.Fields.Item("img_id").Value %>" id="img_id_<%=rs_getImages.Fields.Item("img_id").Value %>">
									<div class="video-container">
										<video preload="metadata" id="video_<%=rs_getImages.Fields.Item("img_id").Value %>" class="video-player" controls disablepictureinpicture controlslist="nodownload"><source src="https://videos.bodyartforms.com/<%=rs_getImages.Fields.Item("img_full").Value%>#t=0.1" type="video/mp4"></video>
									</div>
								</a>	
								<%else%>
									<a class="position-relative pointer" data-fancybox="product-images" data-caption="<%= var_img_title %>" href="https://bodyartforms-products.bodyartforms.com/<%=(rs_getImages.Fields.Item("img_full").Value)%>" id="img_id_<%=rs_getImages.Fields.Item("img_id").Value %>">
										<span class="position-absolute badge badge-secondary p-2 rounded-0" style="top:0;left:0;z-index:200"><i class="fa fa-search fa-lg"></i></span>
											<img class="img-fluid lazyload" src="/images/image-placeholder.png" data-src="https://bodyartforms-products.bodyartforms.com/<%=(rs_getImages.Fields.Item("img_full").Value)%>" alt="<%= var_img_title %>"  style="max-height: 550px;" />
										<div class="text-center small text-dark"><%= var_img_title %></div>
									</a>
								<%end if%>	
								<% 
								end if ' only display if item is active
								rs_getImages.MoveNext()
								wend ' if images are not found 

								
								while not rsColorCharts.eof %>
								<%
									if instr(rsProduct.Fields.Item("ColorChart").Value, rsColorCharts.Fields.Item("chart_filename").Value) then
										var_chart_exists = "yes"
									%>
									<a class="position-relative pointer" data-fancybox="product-images" data-caption="<%= rsColorCharts.Fields.Item("chart_title").Value %>" href="https://bodyartforms-products.bodyartforms.com/<%= rsColorCharts.Fields.Item("chart_filename").Value %>">
										<span class="position-absolute badge badge-secondary p-2 rounded-0" style="top:0;left:0;z-index:200"><i class="fa fa-search fa-lg"></i></span>
										<img class="img-fluid"  src="https://bodyartforms-products.bodyartforms.com/<%= rsColorCharts.Fields.Item("chart_filename").Value %>" alt="Color Chart" title="<%= rsColorCharts.Fields.Item("chart_title").Value %>"  data-color-chart="yes" data-imgname="<%= rsColorCharts.Fields.Item("chart_filename").Value %>"  style="max-height: 550px" />
									</a>							
									<%
									end if
								
								rsColorCharts.movenext()
								wend
								%>
								</div><!-- end slider-main-image -->
							</div><!-- end col-img-main -->
						</div><!-- end row-list -->
			</div><!-- end product photos column -->
			<div class="col-12 col-lg-6 col-xl-7 bg-lightgrey2 py-3">
							<% if (rsProduct.Fields.Item("SaleDiscount").Value > 0 AND rsProduct.Fields.Item("secret_sale").Value = 0) OR  (rsProduct.Fields.Item("secret_sale").Value = 1 AND session("secret_sale") = "yes") then %>
								<div class="badge badge-danger p-2 my-2 d-inline-block" style="font-size:.9em"><span class="mr-3"><%= rsProduct.Fields.Item("SaleDiscount").Value %>% OFF</span><span class="sale-info"><span class="mr-3">Savings: <%= sale_savings %></span>Retail <s><%= sale_retail_price %></s></span></div>
							<% end if %>
							<form class="needs-validation" name="frm-add-cart" id="frm-add-cart" novalidate>
								<input name="productid" type="hidden" id="productid" value="<%= rsProduct.Fields.Item("ProductID").Value %>">
								<input type="hidden" name="customorder" id="customorder" value="<%= rsProduct.Fields.Item("customorder").Value %>">
								
							<input name="discount_amount" type="hidden" value="<%= rsProduct.Fields.Item("SaleDiscount").Value %>">
							<% If display_add_to_cart = 1 and rsProduct.Fields.Item("Active").Value = 1 Then
							
							if var_totalActiveitems > 10 and total_gauges > 1 and not rsGaugeFilter.eof then %>
							<div class="form-group">
							<select class="form-control" name="filter-gauge" id="filter-gauge">
								<option disabled="disabled" selected="selected">Filter by gauge &nbsp;&nbsp;&nbsp;</option>
								<option value="">Show all gauges</option>
								<option disabled="disabled">&nbsp;</option>
								<% While NOT rsGaugeFilter.EOF %>
								<option value="<%= server.htmlencode(rsGaugeFilter.Fields.Item("Gauge").Value) %>"><%= rsGaugeFilter.Fields.Item("Gauge").Value %></option>
								<% rsGaugeFilter.movenext
								wend
								%>
							</select>
							</div>
							<% end if ' > 20 and not rsGaugeFilter.eof 
							%>
								<span id="loading-addtocart" style="display:none"><i class="fa fa-spinner fa-2x fa-spin"></i></span>
								
							<span class="form-group" id="select-addtocart"><!--#include virtual="/products/inc-details-dropdown-addtocart.asp" --></span>
				<div class="my-2">
							<span class="qty-input-field">
							<div class="input-group">
									<div class="input-group-prepend qty-deduct d-md-none">
									  <div class="input-group-text">
										  <i class="fa fa-minus fa-lg"></i>
									  </div>
								  </div>
								  <div class="input-group-prepend qty-deduct d-none d-md-inline">
										  <div class="input-group-text">
											  Qty
										  </div>
										</div>
									
									<input class="form-control text-center" name="qty" type="tel" id="add-qty" value="<%= var_qty_default %>" maxlength="2" placeholder="1" required>
									<div class="input-group-append qty-add d-md-none">
										  <div class="input-group-text">
											  <i class="fa fa-plus fa-lg"></i>
										  </div>
										</div>
							</div>	
						</span>
								  <span class="alert alert-warning d-inline-block px-2 py-1 ml-2 small">Sold as a <span class="font-weight-bold"><%= var_pair_status %></span></span>		
								</div>
												
							<%
					
							If rsProduct.Fields.Item("customorder").Value = "yes" then
								
								other_fields_found = rsProduct.Fields.Item("preorder_field1").Value & rsProduct.Fields.Item("preorder_field2").Value & rsProduct.Fields.Item("preorder_field3").Value & rsProduct.Fields.Item("preorder_field4").Value & rsProduct.Fields.Item("preorder_field5").Value & rsProduct.Fields.Item("preorder_field6").Value & rsProduct.Fields.Item("preorder_field7").Value
							%>
							<div class="mt-3">
							<% if other_fields_found = "" or IsNull(other_fields_found) then
				
							If rsProduct.Fields.Item("preorder_nospecs").Value = 0 then %>
							<div class="form-group">
								<label for="preorders">Provide pre-order specifications (see product description below for what info we need)</label>
								<textarea class="form-control" name="preorders" id="preorders" placeholder="Type specs here" rows="4"></textarea>
							</div>
							<% end if %>
							<% else 	' other_fields_found = "" %>
							
							<%
							If rsProduct.Fields.Item("preorder_field1").Value <> "" then %>
							<div class="form-group">
								<label for="preorder_field1"><%= Server.HTMLEncode(rsProduct.Fields.Item("preorder_field1").Value) %></label>
								<input class="form-control" type="text" name="preorder_field1" id="preorder_field1" placeholder="Type specs here" required>
								<input type="hidden" name="preorder_field1_label" value="<%= rsProduct.Fields.Item("preorder_field1_label").Value %>">
								<div class="invalid-feedback">
										Pre-Order specifications required
								</div>
							</div>
							<% end if %>
							<%
							If rsProduct.Fields.Item("preorder_field2").Value <> "" then %>
							<div class="form-group">
								<label for="preorder_field2"><%= Server.HTMLEncode(rsProduct.Fields.Item("preorder_field2").Value) %></label>
								<input class="form-control" type="text" name="preorder_field2" id="preorder_field2"  placeholder="Type specs here" required>
								<input type="hidden" name="preorder_field2_label" value="<%= rsProduct.Fields.Item("preorder_field2_label").Value %>">
								<div class="invalid-feedback">
										Pre-Order specifications required
								</div>
							</div>
							<% end if %>
							<%
							If rsProduct.Fields.Item("preorder_field3").Value <> "" then %>
							<div class="form-group">
								<label for="preorder_field3"><%= Server.HTMLEncode(rsProduct.Fields.Item("preorder_field3").Value) %></label>
								<input class="form-control" type="text" name="preorder_field3" id="preorder_field3" placeholder="Type specs here" required>
								<input type="hidden" name="preorder_field3_label" value="<%= rsProduct.Fields.Item("preorder_field3_label").Value %>">
								<div class="invalid-feedback">
										Pre-Order specifications required
								</div>
							</div>	
							<% end if %>
							<%
							If rsProduct.Fields.Item("preorder_field4").Value <> "" then %>
							<div class="form-group">
								<label for="preorder_field4"><%= Server.HTMLEncode(rsProduct.Fields.Item("preorder_field4").Value) %></label>
								<input class="form-control" type="text" name="preorder_field4" id="preorder_field4" placeholder="Type specs here" required>
								<input type="hidden" name="preorder_field4_label" value="<%= rsProduct.Fields.Item("preorder_field4_label").Value %>">
								<div class="invalid-feedback">
										Pre-Order specifications required
								</div>
							</div>
							<% end if %>
							<%
							If rsProduct.Fields.Item("preorder_field5").Value <> "" then %>
							<div class="form-group">	
							<label for="preorder_field5"><%= Server.HTMLEncode(rsProduct.Fields.Item("preorder_field5").Value) %></label>
								<input class="form-control" type="text" name="preorder_field5" id="preorder_field5"  placeholder="Type specs here" required><input type="hidden" name="preorder_field5_label" value="<%= rsProduct.Fields.Item("preorder_field5_label").Value %>">
								<div class="invalid-feedback">
										Pre-Order specifications required
								</div>
							</div>
							<% end if %>
							<%
							If rsProduct.Fields.Item("preorder_field6").Value <> "" then %>
							<div class="form-group">
							<label for="preorder_field6"><%= Server.HTMLEncode(rsProduct.Fields.Item("preorder_field6").Value) %></label>
								<input class="form-control" type="text" name="preorder_field6" id="preorder_field6"  placeholder="Type specs here">
								<input type="hidden" name="preorder_field6_label" value="<%= rsProduct.Fields.Item("preorder_field6_label").Value %>">
							</div>	
							<% end if %>
							<%
							If rsProduct.Fields.Item("preorder_field7").Value <> "" then %>
							<div class="form-group">	
							<label for="preorder_field7"><%= Server.HTMLEncode(rsProduct.Fields.Item("preorder_field7").Value) %></label>
								<input class="form-control" type="text" name="preorder_field7" id="preorder_field7"  placeholder="Type specs here">
								<input type="hidden" name="preorder_field7_label" value="<%= rsProduct.Fields.Item("preorder_field7_label").Value %>">
								</div>
							<% end if %>				
							<% end if 	' 	other_fields_found = "" %>			
							</div>
							<% end if %>
							<button class="btn btn-lg btn-purple my-2 add_to_cart btn-block-mobile px-5" type="button"  id="btn-add-cart"><i class="fa fa-shopping-cart fa-lg mr-3"></i>Add to cart</button>
								<div class="add-cart-message" style="display:none"></div>
				
							
	
				
						

<!--#include virtual="/includes/inc-currency-images.asp" -->

<div class="currency my-2">

	<% if Request.Cookies("ID") <> "" then %>	
	<% if NOT rsGetWishlist.EOF then
	var_wishlist_btn_style = "btn-danger"
else
var_wishlist_btn_style = "btn-outline-danger"
	end if 
	
' === only show afterpay option to USA customers
if request.cookies("currency") = "" OR request.cookies("currency") = "USD" then
	afterpay_display = ""
else
	afterpay_display = "display:none"
end if
	
	%>
	<div id="REMOVE-GO-LIVE" style="display:none">
	<div class="afterpay_option" style="<%= afterpay_display %>">
		<div class="afterpay-widget"></div>
	</div>
</div>
<button class="link-add-wishlist btn btn-sm <%= var_wishlist_btn_style %>" type="button"><i class="fa fa-heart fa-lg"></i> Add to Wishlist</button>
<% end if %>
<span class="select-currency btn btn-sm btn-outline-secondary">
<span class="ajax-currency"><img src="/images/icons/<%= currency_img %>"> <%= currency_text %></span> <i class="fa fa-chevron-down"></i></span>
<% if Request.Cookies("showmm") <> "yes" then %>	
<button class="btn btn-sm btn-outline-secondary" id="show-mm" type="button">Show mm sizes</button>
<button class="btn btn-sm btn-outline-secondary" style="display:none" id="hide-mm" type="button">Hide mm sizes</button>
<% else %>
<button class="btn btn-sm btn-outline-secondary" style="display:none" id="show-mm" type="button">Show mm sizes</button>
<button class="btn btn-sm btn-outline-secondary" id="hide-mm" type="button">Hide mm sizes</button>
<% end if %>
</div>

<% if rsHowManyInCarts("currently_in_all_carts") > 10 then %>
<div class="text-info font-weight-bold">
	<%= rsHowManyInCarts("currently_in_all_carts") %> customers have this in their cart
</div>
<% end if %>


<div class="currency-menu my-2" style="display:none">
<!--#include virtual="/template/inc-currency-menu.asp"-->
</div>
<% if Request.Cookies("ID") <> "" then %>	
							
<div class="my-2">
	<div class="wishlist-toggle" style="display:none">
			<div class="wishlist-message"></div>
	<form name="frm-add-wishlist" id="frm-add-wishlist">
		<div class="form-row">
	<% If Not rsGetCategory.EOF Or Not rsGetCategory.BOF Then %>
	<div class="col-12 col-md-auto">
	<select class="form-control my-2" name="wishlist-category" id="wishlist-category">
	<option value="0" selected>Select list...</option>
	<option value="0">None</option>
	<% 
	While NOT rsGetCategory.EOF 
	%>
	<option value="<%=(rsGetCategory.Fields.Item("WishlistID").Value)%>"><%=(rsGetCategory.Fields.Item("WishlistName").Value)%></option>

	<% 
	rsGetCategory.MoveNext()
	Wend
	%>

	</select>
</div><!-- col -->
	<% End If ' end Not rsGetCategory.EOF  %>
	<div class="col-12 col-md-auto">		  
	<select class="form-control my-2" name="priority" id="wishlist-priority">
	<option value="3" selected>Select priority...</option>
	<option value="1">1 - Must have</option>
	<option value="2">2 - Love to have</option>
	<option value="3">3 - Like to have</option>
	<option value="4">4 - I'm thinking about it</option>
	<option value="5">5 - Don't buy this for me</option>
	</select>
</div>	
</div><!-- end form row -->
	</form>
	</div>
</div><!-- my-4 vertical spacing -->
<% end if ' if customer cookie is found %>

<% if Request.Cookies("ID") <> "" then %>
<% if NOT rsGetWishlist.EOF then %>
<div class="small my-2 alert alert-secondary">
<div class="font-weight-bold">Currently in your wishlist:</div>
	<% 
	While NOT rsGetWishlist.EOF 
	%>
	Qty <%= rsGetWishlist.Fields.Item("desired").Value %> ... <%= rsGetWishlist.Fields.Item("gauge").Value %>&nbsp;<%= rsGetWishlist.Fields.Item("length").Value %>&nbsp;<%= rsGetWishlist.Fields.Item("ProductDetail1").Value %><br/>
	<% 
	rsGetWishlist.MoveNext()
	Wend ' Not rsGetWishlist.EOF
	%>
</div>
<% end if ' if anything currently in wishlist
%>
								<% if NOT rsGetOrderHistory.EOF then %>
								<div class="small my-2 alert alert-secondary">
								<div class="font-weight-bold">Already purchased:</div>
								<% 
								While NOT rsGetOrderHistory.EOF
								%>
								Qty <%=(rsGetOrderHistory.Fields.Item("qty").Value)%> ... <%=(rsGetOrderHistory.Fields.Item("Gauge").Value)%><br>
								<% 
								rsGetOrderHistory.MoveNext()
								Wend
								%>
								</div>
							<%	end if ' NOT rsGetOrderHistory.EOF 
							%>
							<% end if ' request.Cookies("ID") <> "" 
								
							else ' if the item is out of stock or inactive, show notice
							%>
								<div class="alert alert-danger">
								<div class="font-weight-bold">Sorry, but this product is currently out of stock.</div>
								<% if var_regular_stock = "" then %>
								<div class="mt-2">To get on the waiting list, click the <strong>Waiting List</strong> tab below, and then find your size to sign up.
								</div>
								<% else %>
								This item will not be re-stocked.
								<% end if %>
								</div>
							<%
							 end if ' Not rsDisplayDropDown.EOF 
							%>
							</form>
				
			</div><!-- end add to cart column -->
		</div><!-- end row -->
	</div><!-- end fluid container -->
</section>


	
	<section>
		<ul class="nav nav-tabs border-secondary" id="nav-menu">
				<li class="nav-item mr-1">
					<a class="nav-link text-ltpurple border-secondary active" id="description-tab" data-toggle="tab" href="#description" role="tab" aria-controls="description" aria-selected="true">Description</a>
				</li>
				<% If rsProduct.Fields.Item("customorder").Value = "yes" then %>
				<li class="nav-item mr-1">
						<a class="nav-link text-ltpurple border-secondary" id="preordersfaq-tab" data-toggle="tab" href="#preordersfaq" role="tab" aria-controls="preordersfaq" aria-selected="true">Pre-Orders</a>
				</li>
				<% end if %>
				<li class="nav-item mr-1">
					<a class="nav-link text-ltpurple border-secondary" id="sizing-tab" data-toggle="tab" href="#sizing" role="tab" aria-controls="sizing" aria-selected="true">Sizing Help</a>
				</li>
			<% if ShowOutItems = "yes" then %>
				<li class="nav-item">
					<a class="nav-link text-ltpurple border-secondary" id="waiting-tab" data-toggle="tab" href="#waiting" role="tab" aria-controls="waiting" aria-selected="true">Waiting List</a>
				</li>
			<% end if %>
		</ul>
		<!-- Tab panes -->
<div class="tab-content border border-secondary p-2">
		<div class="tab-pane active" style="word-wrap: break-word" id="description" role="tabpanel" aria-labelledby="description-tab">
				<% if var_mini_notice <> "" then %>
				<div class="alert alert-danger my-2">
				<%= var_mini_notice %>
				</div>
				<% end if %>
						<span class="fb-like mr-2" style="display:inline" data-href="http://www.bodyartforms.com/productdetails.asp?ProductID=<%= Request.Querystring("ProductID") %>" data-send="false" data-layout="button_count" data-size="large" data-width="450" data-show-faces="false"></span>
						<!-- https://developers.pinterest.com/docs/widgets/pin-it/#custom -->
							<a data-pin-tall="true" href="http://pinterest.com/pin/create/button/?url=http%3A%2F%2Fwww.bodyartforms.com%2Fproductdetails.asp%3FProductID%3D<%=(rsProduct.Fields.Item("ProductID").Value)%>&media=http%3A%2F%2Fbodyartforms-products.bodyartforms.com%2F<%=(rsProduct.Fields.Item("largepic").Value)%>&description=<%=(rsProduct.Fields.Item("title").Value)%>" ></a>

						<ul class="mt-2">
							<li>
							<span id="productid_num" data-productid="<%= rsProduct.Fields.Item("ProductID").Value %>">
							Product # <span><%= rsProduct.Fields.Item("ProductID").Value %></span></span>		
							</li>
							<%= var_materials %>
							<%= var_flares %>
							<%= var_threading_type %>
							<%= var_worn_in %>
							<%= var_sizes_offered %>
							<%= var_brand_logo %>
							<%= origin_country %>
						</ul>
								
						<div class="main-description">
						<% if rsProduct.Fields.Item("seo_meta_description").Value <> "" then %>
						<%= rsProduct.Fields.Item("seo_meta_description").Value %>
						<br/><br/>
						<% end if %>

						<% if instr(var_sizes_offered, "Threadless") > 0 then %>
						Threadless ends will fit <a href="https://bodyartforms.com/products.asp?jewelry=belly&jewelry=circular&jewelry=curved&jewelry=labret&jewelry=barbell&threading=Threadless" target="_blank">most of these posts</a>.
						<% end if %>
						<% '======== BODY GEMS CUSTOM LINKS INFORMATION =============
						if lcase(rsProduct.Fields.Item("brandname").Value) = "body gems" AND instr(lcase(rsProduct.Fields.Item("jewelry").Value), "balls") AND instr(lcase(var_threading_type), "internally")  then %>
						<br/>
						The Internally threaded ends are made to be worn in only certain internally threaded posts. For internally threaded ends, we can guarantee that the end will work with <a href="/products.asp?jewelry=labret&gauge=18g&gauge=16g&gauge=14g&gauge=12g&brand=body+circle&brand=invictus&brand=le+roi&brand=invictus&threading=Internally+threaded" target="_blank">these posts only</a>.<br><br>
					   <% end if %>	
						<%= rsProduct.Fields.Item("description").Value %>					
						<% 
						'strDbContent = rsProduct.Fields.Item("description").Value
						'strDbContent = HandleIncludeFiles(strDbContent)
						'Exec(strDbContent)
						'Function HandleIncludeFiles(dbContent)
						If InStr(rsProduct.Fields.Item("description").Value, "<!--#") > 0 Then
		
						SUB ReadDisplayFile(FileToRead)
						whichfile=server.mappath(FileToRead)
						Set fs = CreateObject("Scripting.FileSystemObject")
						Set thisfile = fs.OpenTextFile(whichfile, 1, False)
						tempSTR=thisfile.readall
						response.write tempSTR
						thisfile.Close
						set thisfile=nothing
						set fs=nothing
						END SUB
		
						startInclude = InStr(rsProduct.Fields.Item("description").Value, "includes")
						endInclude = InStr(rsProduct.Fields.Item("description").Value, "-->")
						GetDiff = endInclude - startInclude
						Include = Mid(rsProduct.Fields.Item("description").Value,startInclude,GetDiff)
						Call ReadDisplayFile(Include)
		
						End If %>
						</div>
						<% if InStr(rsProduct.Fields.Item("internal").Value, "Threadless") > 0 then %>
						<video class="mw-100" width="560" height="315" preload="metadata" controls muted>
							<source src="https://videos.bodyartforms.com/video-threadless-ends-how-to.mp4#t=0.5" type="video/mp4">
						  Your browser does not support playing embedded videos
						  </video>	
						<% end if %>
						<% if var_product_notice <> "" then %>
						<div class="alert alert-warning p-1 product-notice">
						<span class="font-weight-bold">PLEASE NOTE:</span>
								<%= var_product_notice %>
						</div>
						<% end if %>
						<% if (rsProduct.Fields.Item("type").Value) = "limited" then %> 
							<div class="alert alert-info">THIS IS A LIMITED ITEM!
						Limited jewelry is usually re-stocked much less frequently and is harder for us to get from the manufacturer. So get it while you can, while it's available!</div>
						<% end if %>
						<% if instr(lcase(var_materials), "acrylic") > 0 and instr(lcase(rsProduct.Fields.Item("jewelry").Value), "plugs") > 0 then %>
				
							<div class="alert alert-warning">Items with an acrylic wearable area should be worn in completely healed piercings only.</div>
						<% end if %>
		</div>
		<% If rsProduct.Fields.Item("customorder").Value = "yes" then %>
		<div class="tab-pane" id="preordersfaq" role="tabpanel" aria-labelledby="preordersfaq-tab">
				<!--#include virtual="/misc_pages/inc-preorder-info.asp" -->
		</div>
		<% end if %>
		<div class="tab-pane" id="sizing" role="tabpanel" aria-labelledby="sizing-tab">
			<!--#include virtual="/misc_pages/measurement_help.asp" -->
		</div>
		<div class="tab-pane" id="waiting" role="tabpanel" aria-labelledby="waiting-tab">
				<div class="card bg-light my-3 frm-add-waiting" style="display:none">
					<div class="card-body p-2">
					<form  name="frm-add-waiting" id="frm-add-waiting">
							<div class="form-group">	
						<label for="waiting-qty">Qty:</label><input class="ml-1 form-control form-control-sm" style="width:40px" name="waiting-qty" id="waiting-qty" type="tel" value="1" placeholder="1"> <% if var_pair_status = "PAIR" then %><%= var_pair_status %><% end if %>
							</div>
							<div class="form-group mb-3">
							<label for="waiting-email">E-mail</label>
							<input class="form-control form-control-sm" name="waiting-email" id="waiting-email" type="email" value="<%= var_customer_email %>" placeholder="Your email">
						</div>
						<button class="btn btn-sm btn-purple btn-add-waiting" type="button" data-detailid="">Notify me when back in stock</button>
					</form>
				</div><!-- card body -->
				</div><!-- end card -->
					<% if ShowOutItems = "yes" then %>
					<% if var_regular_stock <> "" then %>
						<div class="alert alert-danger font-weight-bold">
						The items below will not be re-stocked.
						</div>
					<% end if ' if not regular stock %>
					
					<% While NOT rs_getDropDownItems.EOF
					%>
					<div class="mb-4 my-md-2">
					<%
					 if (rs_getDropDownItems.Fields.Item("qty").Value <= "0" AND rsProduct.Fields.Item("pair").Value = "yes") OR (rs_getDropDownItems.Fields.Item("qty").Value <= 1 AND (rsProduct.Fields.Item("pair").Value <> "yes" or isNull(rsProduct.Fields.Item("pair").Value))) then
					%>
					
						<% if var_regular_stock = "" then %>
						<span class="link-waiting" data-detailid="<%= rs_getDropDownItems.Fields.Item("ProductDetailID").Value %>">
							<button class="btn btn-sm btn-outline-secondary" type="button"><i class="fa fa-email fa-lg"></i> Add to waiting list</button>
						</span>
						<% end if ' if item is regular stock %>
					<% if Request.Cookies("ID") <> "" then %>
					<% if var_regular_stock = "" then %>
					<span class="link-wishlist" data-detailid="<%= rs_getDropDownItems.Fields.Item("ProductDetailID").Value %>" data-productid="<%= rsProduct.Fields.Item("productid").Value %>">
						<button class="btn btn-sm btn-outline-danger add-wishlist-<%= rs_getDropDownItems.Fields.Item("ProductDetailID").Value %>" type="button"><i class="fa fa-heart fa-lg"></i> Add to Wishlist</button>
					</span>
					<% end if ' if item is regular stock %>
					<% end if %>
					<span class="d-block d-md-inline">
						<%= rs_getDropDownItems.Fields.Item("OptionTitle").Value %> - <%= exchange_symbol %><%= FormatNumber(rs_getDropDownItems.Fields.Item("price").Value * exchange_rate,2)%>
					</span>
					<div class="add-waiting-<%= rs_getDropDownItems.Fields.Item("ProductDetailID").Value %>"></div>
					<div class="message-outs-<%= rs_getDropDownItems.Fields.Item("ProductDetailID").Value %>"></div>
					<% end if %>
					
					</div><!-- item margins -->
					<%
					rs_getDropDownItems.MoveNext()
					Wend
					rs_getDropDownItems.Requery()
					 end if %>
					
		</div>
	  </div>
	</section>	

<% If Not rsGetPhotos.EOF Then %>
<a name="photos"></a>	
	<div class="display-5 mt-4 mb-2">Customer Photos
	</div>
	<form class="form-inline my-3">
		<select class="form-control mt-1" id="filter_photos_gauge" name="filter_photos_gauge" data-filter="gauge"  data-replace="filter_photos_color" data-type="photos">
		<option value="All">Filter photos by gauge</option>
		<option value="All">Show all</option>
		<optgroup label=" "></optgroup>
		<% while not rs_photos_gauge_dropdown.eof %>
			<option value="<%= Server.URLEncode(rs_photos_gauge_dropdown.Fields.Item("gauge").Value) %>">
				<%=rs_photos_gauge_dropdown.Fields.Item("gauge").Value%>&nbsp;&nbsp;&nbsp;(<%=rs_photos_gauge_dropdown.Fields.Item("gauge_total").Value%> photos)
			</option>
		<% rs_photos_gauge_dropdown.movenext() 
		wend
		%>
		</select>

		<select class="form-control mt-1 ml-0 ml-sm-3" id="filter_photos_color" name="filter_photos_color" data-filter="color" data-type="photos" data-replace="filter_photos_gauge">
			<option value="All">Filter photos by color / style</option>
			<option value="All">Show all</option>
			<optgroup label=" "></optgroup>
			
			<% while not rs_photos_color_dropdown.eof %>
				<option value="<%= Server.URLEncode(rs_photos_color_dropdown.Fields.Item("color").Value) %>">
					<%=rs_photos_color_dropdown.Fields.Item("color").Value%>&nbsp;&nbsp;&nbsp;(<%=rs_photos_color_dropdown.Fields.Item("color_total").Value%> photos)
				</option>
			<% rs_photos_color_dropdown.movenext() 
			wend
			%>
			</select>
	</form>
	<div id="customer-photo-loader"></div>
	<% End If ' Not rsGetPhotos.EOF %>

	<% If Not rsCrossSellingItems.EOF Then %>
	<div class="display-5 mt-4 mb-2">Customers Also Bought
	</div>
	<div class="baf-carousel" id="cross-selling">
	<% 	While NOT rsCrossSellingItems.EOF %>
	<div class="slide">
		<a href="productdetails.asp?ProductID=<%= rsCrossSellingItems("bought_with").Value %>">
			<img class="img-fluid lazyload CrossSellingItems" src="/images/image-placeholder.png" data-src="https://bafthumbs-400.bodyartforms.com/<%=(rsCrossSellingItems("picture").Value)%>" alt="<%=(rsCrossSellingItems("title").Value)%>" />
		</a>
	</div>
	<% rsCrossSellingItems.MoveNext()
	Wend
	%>
	</div>
	<% end if 'rsCrossSellingItems.EOF %>	


	<% If total_reviews > 0 or var_show_star_ratings = "yes" Then %>
		<section class="mb-5">
		<a name="reviews"></a>
			<div class="display-5 mt-5 mb-2">Customer Reviews (<%= total_reviews %>)</div>
			<% 
			if var_show_star_ratings = "yes" then %>
			
				<% if var_show_star_ratings = "yes" then %>
				<span class="rating-box">
						<span class="rating" style="width:<%= var_avg_percentage %>%"></span>
					</span>
				<span class="text-avg">
					<% if avg_rating <> "" then %>
						<%= avg_rating %> out of 5 stars
					<% end if %>
				</span>
					
				<% end if %>
					
			<%
			i = 5
			while not rs_StarCounts.eof 
			var_star_percentage = round((100 / total_ratings) * rs_StarCounts.Fields.Item("counts").Value)
			%> 
			<div class="small my-1">
				<div class="progress bg-lightgrey2 d-inline-flex align-text-top pointer hover-rating" style="width: 150px"  data-rating="<%= rs_StarCounts.Fields.Item("review_rating").Value %>">
						<div class="progress-bar" role="progressbar" aria-valuenow="<%= var_star_percentage %>"
						aria-valuemin="0" aria-valuemax="100" style="width:<%= var_star_percentage %>%">
						</div>
				</div> 				<% l = 0
				do until l = i %>
				   <i class="fa fa-star text-warning"></i>
			   <% l = l + 1 
			   loop %>
			    <%= var_star_percentage %>% 
				(<%= rs_StarCounts.Fields.Item("counts").Value %>)
			</div>
			<%
			i = i - 1
			rs_StarCounts.movenext()
			wend

			end if ' var_show_star_ratings = "yes"
			
			if total_reviews > 0 then %>
			<form class="form-inline my-3">
				<select class="form-control mt-1" id="filter_review_gauge" name="filter_review_gauge" data-filter="gauge"  data-replace="filter_review_color" data-type="reviews">
				<option value="">Filter reviews by gauge</option>
				<option value="">Show all</option>
				<optgroup label=" "></optgroup>
				<% while not rs_reviews_gauge_dropdown.eof %>
					<option value="<%= Server.URLEncode(rs_reviews_gauge_dropdown.Fields.Item("gauge").Value) %>">
						<%=rs_reviews_gauge_dropdown.Fields.Item("gauge").Value%>&nbsp;&nbsp;&nbsp;(<%=rs_reviews_gauge_dropdown.Fields.Item("gauge_total").Value%> reviews)
					</option>
				<% rs_reviews_gauge_dropdown.movenext() 
				wend
				%>
				</select>

				<select class="form-control mt-1 ml-0 ml-sm-3" id="filter_review_color" name="filter_review_color" data-filter="color" data-type="reviews" data-replace="filter_review_gauge">
					<option value="">Filter reviews by color / style</option>
					<option value="">Show all</option>
					<optgroup label=" "></optgroup>
					<% while not rs_reviews_color_dropdown.eof %>
						<option value="<%= Server.URLEncode(rs_reviews_color_dropdown.Fields.Item("color").Value) %>">
							<%=rs_reviews_color_dropdown.Fields.Item("color").Value%>&nbsp;&nbsp;&nbsp;(<%=rs_reviews_color_dropdown.Fields.Item("color_total").Value%> reviews)
						</option>
					<% rs_reviews_color_dropdown.movenext() 
					wend
					%>
					</select>
			</form>
			<% end if ' if there are text reviews %>
			<div id="div_product_reviews"></div>
		
		</section>
	<% end if ' total_reviews > 0 %>


	<% If Not rsRecentlyViewed.EOF Then %>
	<div class="display-5 mt-4 mb-2">Recently Viewed
	</div>
	<div class="baf-carousel" id="recents">
	<% 	While NOT rsRecentlyViewed.EOF %>
	<div class="slide">
		<a href="productdetails.asp?ProductID=<%= rsRecentlyViewed.Fields.Item("ProductID").Value %>">
			<img class="img-fluid lazyload" src="/images/image-placeholder.png" data-src="https://bafthumbs-400.bodyartforms.com/<%=(rsRecentlyViewed.Fields.Item("picture").Value)%>" alt="<%=(rsRecentlyViewed.Fields.Item("title").Value)%>" />
		</a>
	</div>
	<% rsRecentlyViewed.MoveNext()
	Wend
	%>
	</div>
	<% end if 'rsRecentlyViewed.EOF %>	
		

	<% end if ' only display if product is active %>
	
	<!-- BEGIN REPORT PHOTO MODAL WINDOW -->
	<div class="modal fade" id="modal-report-photo" tabindex="-1" role="dialog"  aria-labelledby="modal-report-photo" style="z-index:999999">
		<div class="modal-dialog" role="document">
			<div class="modal-content">
					<div class="modal-header">
						<h5 class="modal-title" id="report-photo-label"></h5>
						<button type="button" class="close" data-dismiss="modal" aria-label="Close">
								<span aria-hidden="true">&times;</span>
						</button>
					</div>
					<div class="modal-body">
						<form class="needs-validation" name="frmReportPhoto" id="frmReportPhoto" novalidate>
							<div class="form-group">Briefly describe the issue with the photo</div>
							<div class="form-group">
								<input class="form-control" type="text" name="report-photo-comments" id="report-photo-comments" placeholder="Comments" required>
							</div>
							<div id="report-photo-message"></div>
							<div class="text-center">
								<button type="submit" name="btn-report" class="btn btn-block btn-purple">Report</button>
							</div>
						</form>
					</div>
			</div>
		</div>
	</div>
	<!-- END REPORT PHOTO MODAL WINDOW -->	
	
<!--#include virtual="/bootstrap-template/footer.asp" -->

<script type="text/javascript" src="/js-pages/currency-exchange.min.js?v=022420"></script>
<% if (session("exchange-rate") = "" OR session("exchange-currency") <> request.cookies("currency")) AND request.cookies("currency") <> "" AND request.cookies("currency") <> "USD" then %>
<script>
		// Get currency conversions on page load
		updateCurrency();
</script>
<% end if %>
<script src="/js/jquery.fancybox.min.js"></script>
<script type="text/javascript" src="/js/slick.min.js"></script>
<script type="text/javascript" src="/js-pages/product-details.min.js?v=101821" ></script>

<!-- Start Afterpay Javascript -->
<!--
<script type = "text/javascript" src="https://static-us.afterpay.com/javascript/present-afterpay.js"></script>-->
<!--
<script type="text/javascript" src="/js-pages/afterpay-widget.js?v=020420" ></script>-->
<!-- Pinterest -->
<script async defer src="//assets.pinterest.com/js/pinit.js"></script>
<!-- Facebook -->
<script>
	(function(d, s, id) {
	  var js, fjs = d.getElementsByTagName(s)[0];
	  if (d.getElementById(id)) return;
	  js = d.createElement(s); js.id = id;
	  js.src = "//connect.facebook.net/en_US/all.js#xfbml=1&appId=180076978718781";
	  fjs.parentNode.insertBefore(js, fjs);
	}(document, 'script', 'facebook-jssdk'));

</script>

<%
Set rsProduct = Nothing
Set rsProductStats = Nothing
Set rsRecentlyViewed = Nothing
Set rsCrossSellingItems = Nothing
Set rsColorCharts = Nothing
Set rs_getImages = Nothing
Set rs_StarCounts = Nothing
Set rs_getDropDownItems = Nothing
Set rsGetActiveItems = Nothing
Set rsGaugeFilter = Nothing
Set rs_reviews_gauge_dropdown = Nothing
Set rs_reviews_color_dropdown = Nothing
Set rs_photos_gauge_dropdown = Nothing
Set rs_photos_color_dropdown = Nothing
Set rsGetReview = Nothing
Set rsGetPhotos = Nothing
Set rsSizesOffered = Nothing
Set rsGetUser = Nothing
Set rsGetWishlist = Nothing
Set rsGetCategory = Nothing
Set rsGetOrderHistory = Nothing
Set rsJsonReviews = Nothing
%>