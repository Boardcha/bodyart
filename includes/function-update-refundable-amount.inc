<%
Function updateRefundableAmount(var_invoice_number, var_customer_number)
	' RE-CALCULATE CORRECT PRICE FOR BACKORDERED ITEMS AFTER SALE TO REFUND FOR
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP (100) PERCENT sent_items.ID, sent_items.coupon_code, sent_items.combined_tax_rate, TBL_OrderSummary.ErrorReportDate, TBL_OrderSummary.ErrorDescription,  sent_items.ship_code, TBL_OrderSummary.qty, ProductDetails.qty AS 'qty_instock', TBL_OrderSummary.item_price, ProductDetails.ProductDetail1, ProductDetails.location, ProductDetails.Gauge, ProductDetails.Length, jewelry.title, ProductDetails.ProductDetailID, ProductDetails.BinNumber_Detail, TBL_OrderSummary.OrderDetailID, TBL_OrderSummary.ProductID, TBL_OrderSummary.item_problem, TBL_OrderSummary.ErrorQtyMissing,  (jewelry.title + ' ' + ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '')) as description FROM sent_items INNER JOIN TBL_OrderSummary ON sent_items.ID = TBL_OrderSummary.InvoiceID INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID WHERE TBL_OrderSummary.backorder = 1 AND ID = ? ORDER BY sent_items.ID"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, var_invoice_number))
	set rsGetItems = Server.CreateObject("ADODB.Recordset")
	rsGetItems.CursorLocation = 3 'adUseClient
	rsGetItems.Open objCmd

	If NOT rsGetItems.EOF Then
		'==============  GET COUPON DISCOUNT / IF ANY ============================================
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT DiscountPercent FROM TBLDiscounts WHERE DiscountCode = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("coupon_code",200,1,50,rsGetItems.Fields.Item("coupon_code").Value))
		Set rsGetCouponDiscount = objCmd.Execute()
	End If

	While NOT rsGetItems.EOF 
		
		If NOT rsGetCouponDiscount.eof then
			var_item_price = FormatNumber((rsGetItems.Fields.Item("item_price").Value - ((rsGetCouponDiscount.Fields.Item("DiscountPercent").Value / 100) * rsGetItems.Fields.Item("item_price").Value)) * rsGetItems.Fields.Item("ErrorQtyMissing").Value, -1, -2, -0, -2)                        
		Else
			var_item_price = FormatNumber(rsGetItems.Fields.Item("item_price").Value * rsGetItems.Fields.Item("qty").Value, -1, -2, -0, -2)
		End if

		' Add on tax to refund 
		If rsGetItems.Fields.Item("combined_tax_rate").Value > 0 then
			var_item_price = var_item_price + (var_item_price * rsGetItems.Fields.Item("combined_tax_rate").Value)
		End if
		var_refund_total = FormatNumber(Ccur(var_refund_total) + ccur(var_item_price), -1, -2, -0, -2)
		rsGetItems.MoveNext
	Wend

	If var_refund_total > 0 then

		Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
		password = "3uBRUbrat77V"
		data = var_invoice_number & "|" & var_customer_number
		encrypted_code = objCrypt.Encrypt(password, data)

		'Check if there is refundable amount is set for the invoice
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM TBL_Refunds_backordered_items WHERE invoice_id = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12, var_invoice_number))
		set rsRefunds = Server.CreateObject("ADODB.Recordset")
		rsRefunds.CursorLocation = 3
		rsRefunds.Open objCmd	
		If Not rsRefunds.EOF Then
			refundableAmountFound = true
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "UPDATE TBL_Refunds_backordered_items SET refund_total = ? WHERE invoice_id = ?"
			objCmd.Parameters.Append(objCmd.CreateParameter("refund_total",6,1,20, var_refund_total))
			objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, var_invoice_number))
			objCmd.Execute()		
		Else
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO TBL_Refunds_backordered_items (invoice_id, refund_total, encrypted_code) VALUES (?,?,?)"
			objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, var_invoice_number))
			objCmd.Parameters.Append(objCmd.CreateParameter("refund_total",6,1,20, var_refund_total))
			objCmd.Parameters.Append(objCmd.CreateParameter("encrypted_code",200,1,250, encrypted_code))
			objCmd.Execute()
		End If
		
		Set objCrypt = Nothing
		Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
		password = "3uBRUbrat77V"
		data = encrypted_code
		decrypted = objCrypt.Decrypt(password, data)
		Set objCrypt = Nothing
	Else
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "DELETE FROM TBL_Refunds_backordered_items WHERE invoice_id = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, var_invoice_number))
		objCmd.Execute()		
	End if
End Function
%>