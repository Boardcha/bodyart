<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/Connections/authnet.asp"-->
<%
	' decrypt refund information
	Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
	password = "3uBRUbrat77V"
	data = request.querystring("encrypted")
	decrypted_refund = objCrypt.Decrypt(password, data)
	
	split_refund = split(decrypted_refund, "|")

	invoice_id = split_refund(0)
    refund_total = split_refund(1)
	var_customer_number = split_refund(2)
    var_refund_id = request.querystring("id")

	Set objCrypt = Nothing

	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TBL_Refunds_backordered_items.*, sent_items.customer_id, email, customer_first, transactionID, pay_method FROM sent_items INNER JOIN TBL_Refunds_backordered_items ON sent_items.ID = TBL_Refunds_backordered_items.invoice_id WHERE date_redeemed= NULL AND invoice_id = ? AND refund_total = ? AND encrypted_code = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, invoice_id))
	objCmd.Parameters.Append(objCmd.CreateParameter("refund_total",6,1,20, refund_total))
	objCmd.Parameters.Append(objCmd.CreateParameter("encrypted_code",200,1,200, data))
	set rsCheckRefund = objCmd.Execute()


    if not rsCheckRefund.eof then
		' ====== Save it as store credit =======
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE customers SET credits = credits + " & refund_total & " WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("customerid",3,1,12, rsCheckRefund.Fields.Item("customer_ID").Value))
		objCmd.Execute()
						
		' ====== update the record to clear it out it, so they can not refund multiple times or cannot use the both option =======
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE TBL_Refunds_backordered_items date_redeemed = GETDATE() WHERE invoice_id = ? AND encrypted_code = ? AND id = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15, invoice_id))
		objCmd.Parameters.Append(objCmd.CreateParameter("encrypted_code",200,1,200, data))
		objCmd.Parameters.Append(objCmd.CreateParameter("var_refund_id",3,1,15, var_refund_id))
		objCmd.Execute()

		mailer_type = "customer_submitted_refund_notification"
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
