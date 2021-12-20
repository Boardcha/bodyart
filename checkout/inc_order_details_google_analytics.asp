<%
	if Session("invoiceid") <> "" then 
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT ID, email, customer_first, customer_last, phone, city, address, billing_address, billing_zip, state, zip, country, shipping_rate, total_sales_tax, total_preferred_discount, total_coupon_discount, total_free_credits, total_store_credit, total_gift_cert, coupon_code FROM sent_items where ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,15,Session("invoiceid")))
	set rsGetOrder = objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT  ISNULL(jewelry.title, '') as 'title', ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '') AS 'variant',  TBL_OrderSummary.DetailID, TBL_OrderSummary.qty, TBL_OrderSummary.item_price, jewelry.largepic, jewelry.brandname, jewelry.jewelry, jewelry.ProductID FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID INNER JOIN TBL_OrderSummary ON ProductDetails.ProductDetailID = TBL_OrderSummary.DetailID WHERE (TBL_OrderSummary.InvoiceID = ?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,15,Session("invoiceid")))

	set rsGoogle_GetOrderDetails = Server.CreateObject("ADODB.Recordset")
	rsGoogle_GetOrderDetails.CursorLocation = 3 'adUseClient
	rsGoogle_GetOrderDetails.Open objCmd
	
	end if 
%>