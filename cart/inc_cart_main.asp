<!--#include virtual="cart/generate_guest_id.asp"-->
<%
' Remove coupon for bug testing
if request.querystring("remove_coupon") = "yes" then
			Session("CouponCode") = ""
			Session("CouponPercentage") = ""
			session("brand_coupon") = ""
			session("GiftCertCode") = ""			
			Session("GiftCertAmount") = 0
			Session("GiftCertID") = 0
			session("usecredit") = ""
			session("storeCredit_amount") = 0
			session("storeCredit_used") = 0
			session("textCouponBox") = ""
			
end if

if request.cookies("OrderAddonsActive") <> "" then
	var_addons = " AND cart_addon_item = 1 "
else
	var_addons = " AND cart_addon_item = 0 "
end if

	Session("StoreCreditAmount") = 0
	Session("StoreCreditID") = 0

	' RETRIEVE SHOPPING CART CONTENTS ---------------------------------------------
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT tbl_carts.cart_id, tbl_carts.cart_detailID, tbl_carts.cart_preorderNotes, tbl_carts.cart_custId, tbl_carts.cart_qty, tbl_carts.cart_wishlistid, jewelry.title, jewelry.internal, jewelry.customorder, jewelry.autoclavable, jewelry.brandname, jewelry.SaleDiscount, jewelry.secret_sale, jewelry.ProductID, jewelry.pair, jewelry.picture, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.price, ProductDetails.wlsl_price, ProductDetails.qty, ProductDetails.ProductDetail1, ProductDetails.ProductDetailID, ProductDetails.free, jewelry.SaleExempt, jewelry.jewelry, TBL_Companies.preorder_timeframes, ProductDetails.weight, (jewelry.title + ' ' + ISNULL(ProductDetails.Gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' +  ISNULL(ProductDetails.ProductDetail1,'')) AS 'mini_preview_text', (ISNULL(ProductDetails.Gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' +  ISNULL(ProductDetails.ProductDetail1,'')) AS 'variant', CAST(ProductDetails.price * tbl_carts.cart_qty as money) AS 'mini_line_price' FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN tbl_carts ON ProductDetails.ProductDetailID = tbl_carts.cart_detailId INNER JOIN TBL_Companies ON jewelry.brandname = TBL_Companies.name WHERE (tbl_carts." & var_db_field & " = ?) AND cart_save_for_later = 0 " & var_addons & " AND ProductDetails.active = 1 AND jewelry.active = 1"
	objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10,var_cart_userid))
	Set rs_getCart = objCmd.Execute()

	if NOT rs_getCart.eof then
	
		'Update last viewed date on all items
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE tbl_carts SET cart_lastViewed = '" & now() & "' WHERE " & var_db_field & " = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10,var_cart_userid))
		objCmd.Execute()
	
	end if

	'=========== SET VARIABLE IF CART IS EMPTY, Set cart COOKIE TOTAL AND check STOCK LEVELS ======================
	If Not rs_getCart.EOF Or Not rs_getCart.BOF Then
			var_cart_count = 0
			While Not rs_getCart.EOF
			
				var_cart_count = var_cart_count + rs_getCart.Fields.Item("cart_qty").Value
			
			rs_getCart.MoveNext()
			Wend
			
			
			Response.Cookies("cartCount") = var_cart_count
			Response.Cookies("cartCount").Expires = DATE + 300
			cart_status = "not-empty"
			

	else ' if cart is empty
		Response.Cookies("cartCount") = 0
		Response.Cookies("cartCount").Expires = DATE + 300
		cart_status = "empty"

	end if
	' ---- End Set cart COOKIE TOTAL and check stock levels ------------------------

	
if Request.Cookies("ID") <> "" then ' if customer is logged in
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM customers  WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
	Set rsGetUser = objCmd.Execute()


	If Not rsGetUser.EOF Or Not rsGetUser.BOF Then

	TotalCredits = rsGetUser.Fields.Item("credits").Value

		IF (rsGetUser.Fields.Item("Flagged").Value) = "Y" then
		Flagged = "yes"
		else
		Flagged = ""
		end if 
	End if
	
	'Set session for credit that is to be used
	if request.querystring("usecredit") = "yes" then
		session("usecredit") = "yes"
		session("storeCredit_amount") = TotalCredits
	end if
	
	' FIND OUT IF CUSTOMER IS PREFERRED OR NOT ----------------------------
	' -----------  LAUNCHED GRANDFATHERED DATE ON 1/7/2019 -----------
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT SUM(dbo.TBL_OrderSummary.item_price * dbo.TBL_OrderSummary.qty) AS totalspent FROM dbo.sent_items INNER JOIN dbo.TBL_OrderSummary ON dbo.sent_items.ID = dbo.TBL_OrderSummary.InvoiceID INNER JOIN dbo.customers ON dbo.sent_items.customer_ID = dbo.customers.customer_ID WHERE dbo.sent_items.customer_ID = ? AND 	dbo.sent_items.ship_code = N'paid' AND grandfathered_discount = 1 AND dbo.TBL_OrderSummary.qty > 0 AND DetailID NOT IN (23998, 29758) HAVING (SUM(dbo.TBL_OrderSummary.item_price * dbo.TBL_OrderSummary.qty) > 0)"

		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
		Set rsGetTotalItems = objCmd.Execute()

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "select SUM(total_returns) AS total_returns FROM sent_items WHERE customer_ID = ? AND ship_code = N'paid'"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
		Set rsGetTotalReturns = objCmd.Execute()

		var_TotalReturns = 0
		If NOT rsGetTotalReturns.EOF And NOT rsGetTotalReturns.BOF Then
			var_TotalReturns = rsGetTotalReturns.Fields.Item("total_returns").Value
		end if
	
		If NOT rsGetTotalItems.EOF And NOT rsGetTotalItems.BOF Then
		  TotalSpent = rsGetTotalItems.Fields.Item("totalspent").Value - var_TotalReturns
		  Session("Preferred") = "yes"
		Else
			TotalSpent = 0
			Session("Preferred") = ""
		End if
		rsGetTotalItems.Close()
		Set rsGetTotalItems = Nothing

	' END FIND OUT IF CUSTOMER IS PREFERRED -------------------------
end if ' if customer is registered -----------------------
	

'=================================================================
' Put this variable in place because 1) to not run uneccesarry code and speed page up and 2) with rsGetFree near bottom of page, it was causing query timeouts. Could never figure out why :/

if var_process_order <> "yes" then ' variable set on checkout_ajax_process_payment.asp page =============================

if check_stock = "yes" then
	var_cart_count = 0
	rs_getCart.ReQuery() 
%>
	<!--#include virtual="cart/inc_cart_stock_check.asp"-->
<%
end if	

' ------- Get FREE items
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT jewelry.title, jewelry.picture, ProductDetails.ProductDetail1, ProductDetails.qty, ProductDetails.free, jewelry.ProductID, ProductDetails.ProductDetailID, ProductDetails.Free_QTY, ProductDetails.weight, jewelry.picture, ProductDetails.price, ProductDetails.active,  ProductDetails.Gauge, ProductDetails.Length, ProductDetails.detail_code, ISNULL(ProductDetails.gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' + ISNULL(ProductDetails.ProductDetail1,'') + ' ' + ISNULL(jewelry.title,'') AS 'free_title' FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID WHERE (jewelry.ProductID <> 3704) AND (ProductDetails.qty > 0) AND (ProductDetails.free <> 0) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.active = 1) ORDER BY CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END, ProductDetails.free DESC, jewelry.title"
		Set rsGetFree = objCmd.Execute()
		
' ------- End getting free items

end if ' if var_process_order <> "yes"

' ------- Get FREE o-rings
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT jewelry.ProductID, jewelry.title, ProductDetails.ProductDetailID, ProductDetails.qty, ProductDetails.ProductDetail1, ProductDetails.Gauge, jewelry.picture, ProductDetails.detail_code, ProductDetails.Free_QTY, ProductDetails.free FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID WHERE (jewelry.ProductID = 530 OR (jewelry.ProductID = 1649 AND ProductDetails.Gauge <> '1-1/8" & """" & "') OR jewelry.ProductID = 15385) AND ProductDetails.qty > 0 ORDER BY title, item_order ASC"
		Set rsGetOrings = objCmd.Execute()

' -------- END getting free o-rings

%>