<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/Connections/authnet.asp"-->
<%
	' decrypt refund information
	Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
	password = "3uBRUbrat77V"
	data = request.querystring("encrypted")
	data = Replace(data, " ", "+") 'Bug fix: IIS converts "+" signs to spaces. We need to convert it back.
	decrypted_refund = objCrypt.Decrypt(password, data)
	
	split_refund = split(decrypted_refund, "|")

	invoice_id = split_refund(0)
    ProductDetailID = split_refund(1)
	var_customer_number = split_refund(2)
    var_refund_id = request.querystring("id")

	Set objCrypt = Nothing

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT REF.*, sent_items.customer_id, email, customer_first, transactionID, pay_method FROM sent_items INNER JOIN TBL_Refunds_backordered_items REF ON sent_items.ID = REF.invoice_id WHERE redeemed = 0 AND REF.invoice_id = ? AND REF.ProductDetailID = ? AND REF.encrypted_code = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, invoice_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductDetailID",3,1,15, ProductDetailID))
	objCmd.Parameters.Append(objCmd.CreateParameter("encrypted_code",200,1,200, data))
	set rsCheckRefund = objCmd.Execute()


    if not rsCheckRefund.eof then
		var_db_refund_amt = formatnumber(rsCheckRefund.Fields.Item("refund_total").Value,2)
		' ====== Save it as store credit =======
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE customers SET credits = credits + " & var_db_refund_amt & " WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("customerid",3,1,12, rsCheckRefund.Fields.Item("customer_ID").Value))
		objCmd.Execute()
						
		' ====== update the record to clear it out, so they can not refund multiple times or cannot use the both option =======
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE TBL_Refunds_backordered_items SET redeemed = 1, date_redeemed = GETDATE(), redeemedAs = 'Store Credit' WHERE invoice_id = ? AND encrypted_code = ? AND id = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, invoice_id))
		objCmd.Parameters.Append(objCmd.CreateParameter("encrypted_code",200,1,200, data))
		objCmd.Parameters.Append(objCmd.CreateParameter("var_refund_id",3,1,15, var_refund_id))
		objCmd.Execute()

		' ==== Set item's backorder status ====
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE TBL_OrderSummary SET backorder = 0, BackorderReview = 'N', notes = 'Customer has refunded the item and backorder has been cleared' WHERE InvoiceID = ? AND DetailID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,15, invoice_id))
		objCmd.Parameters.Append(objCmd.CreateParameter("detailID",3,1,20, ProductDetailID))
		objCmd.Execute()
			
		mailer_type = "customer_submitted_refund_as_store_credit_notification"
		%>
		<!--#include virtual="emails/email_variables.asp"-->
		<%
	
		var_notes = "Automated message: A store credit has been issued by customer"

        ' ========= Notes for original order =========================================== 
        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "INSERT INTO tbl_invoice_notes (user_id, invoice_id, note) VALUES (28,?,?)"
        objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15, invoice_id))
        objCmd.Parameters.Append(objCmd.CreateParameter("note",200,1,1500, var_notes))
        objCmd.Execute()


    end if ' ====== if a record is found
	

	

DataConn.Close()
Set DataConn = Nothing
%>
