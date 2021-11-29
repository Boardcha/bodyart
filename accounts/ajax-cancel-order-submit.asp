<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
invoiceid = request.form("invoice")

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM sent_items WHERE ID = ? AND ship_code = 'paid'"
	objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,invoiceid))
	Set rsGetOrder = objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT sent_items.shipping_rate - sent_items.total_preferred_discount - sent_items.total_coupon_discount - sent_items.total_free_credits - sent_items.total_returns + sent_items.total_sales_tax + sent_items.total_store_credit + sent_items.total_gift_cert AS total_discount_taxes, SUM(TBL_OrderSummary.qty * TBL_OrderSummary.item_price) AS subtotal FROM sent_items INNER JOIN TBL_OrderSummary ON sent_items.ID = TBL_OrderSummary.InvoiceID WHERE (sent_items.ID = ?) GROUP BY sent_items.shipping_rate - sent_items.total_preferred_discount - sent_items.total_coupon_discount - sent_items.total_free_credits - sent_items.total_returns + sent_items.total_sales_tax + sent_items.total_store_credit + sent_items.total_gift_cert"
	objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,invoiceid))
	Set rsGetOrderTotal = objCmd.Execute()
	
	if not rsGetOrderTotal.eof then
		var_order_total = rsGetOrderTotal.Fields.Item("subtotal").Value + rsGetOrderTotal.Fields.Item("total_discount_taxes").Value
	end if

if not rsGetOrder.eof then	

		' -------------  Set order to cancelled status
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE sent_items SET shipped = 'Cancelled', ship_code = 'not paid' WHERE ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10, invoiceid))
		objCmd.Execute()
		
		' -------------  Insert notes about order
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (?,?,?)"
		objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,10,1))
		objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10,invoiceid))
		objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,250,"Automated message - Customer cancelled order via website cancellation. " & FormatCurrency(var_order_total, -1, -2, -2, -2) & " applied to store credit."))
		objCmd.Execute()

		' -------------  Issue store credit
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE customers SET credits = credits + ? WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("credit",6,1,10,FormatNumber(var_order_total, -1, -2, -2, -2)))
		objCmd.Parameters.Append(objCmd.CreateParameter("customerid",3,1,10, CustID_Cookie))
		objCmd.Execute()
		
		
		' Put items back in stock and reactivate all listings and items
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT qty, DetailID, InvoiceID, title, ProductDetail1 FROM QRY_OrderDetails  WHERE ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10,invoiceid))
		set rsUpdate = objCmd.Execute()

		While NOT rsUpdate.EOF

				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE ProductDetails SET qty = qty + " & rsUpdate.Fields.Item("qty").Value & ", active = 1 WHERE ProductDetailID = " & rsUpdate.Fields.Item("DetailID").Value
				objCmd.Execute()
				
				' Reactive main product listing
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn  
				objCmd.CommandText = "UPDATE jewelry SET active = 1 FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID WHERE ProductDetailID = " & rsUpdate.Fields.Item("DetailID").Value
				objCmd.Execute()

				'Write info to edits log	
				set objCmd = Server.CreateObject("ADODB.Command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, detail_id, description, edit_date) VALUES (28, " & rsUpdate("DetailID") & ",'Automated - Added " & rsUpdate("qty") & " to qty from customer cancelling order on front end','" & now() & "')"
				objCmd.Execute()
				Set objCmd = Nothing

			rsUpdate.MoveNext()
		Wend
		
mailer_type = "cancelled_order"
%>
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/emails/email_variables.asp"-->
<%
		
end if ' not rsGetOrder.eof
DataConn.Close()
%>
