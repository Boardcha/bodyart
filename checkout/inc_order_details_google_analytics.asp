<%
	if Session("invoiceid") <> "" then 
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT city, state, country, shipping_rate, total_sales_tax, total_preferred_discount, total_coupon_discount, total_free_credits, total_store_credit, total_gift_cert, coupon_code FROM sent_items where ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,15,Session("invoiceid")))
	set rsGetOrder = objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT  ISNULL(jewelry.title, '') as 'title', ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '') AS 'variant',  TBL_OrderSummary.DetailID, TBL_OrderSummary.qty, TBL_OrderSummary.item_price, jewelry.brandname, jewelry.jewelry, jewelry.ProductID FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID INNER JOIN TBL_OrderSummary ON ProductDetails.ProductDetailID = TBL_OrderSummary.DetailID WHERE (TBL_OrderSummary.InvoiceID = ?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,15,Session("invoiceid")))
	set rsGoogle_GetOrderDetails = objCmd.Execute()
	
	end if 
%>