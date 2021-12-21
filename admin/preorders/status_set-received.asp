<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' If we've received the item
if request.querystring("received") = "yes" then

	  set objCmd = Server.CreateObject("ADODB.command")
	  objCmd.ActiveConnection = DataConn
	  objCmd.CommandText = "UPDATE TBL_OrderSummary SET item_received = 1, item_received_date = GETDATE() ,backorder = 0 WHERE OrderDetailID = ?"
	  objCmd.Parameters.Append(objCmd.CreateParameter("OrderDetailID",3,1,10,request.querystring("id")))
	  objCmd.Execute()
	  
	  
	  set objCmd = Server.CreateObject("ADODB.command")
	  objCmd.ActiveConnection = DataConn
	  objCmd.CommandText = "SELECT TBL_OrderSummary.InvoiceID, TBL_OrderSummary.OrderDetailID, TBL_OrderSummary.item_received, jewelry.customorder FROM TBL_OrderSummary INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID WHERE (TBL_OrderSummary.item_received = 0) AND (TBL_OrderSummary.InvoiceID = ?) AND (jewelry.customorder = 'yes')"
	  objCmd.Parameters.Append(objCmd.CreateParameter("invoiceID",3,1,10,request.querystring("invoice")))
	  Set rsOrderStatus = objCmd.Execute()
	  
	  ' Set order to pending if no items are found
	  If rsOrderStatus.BOF AND rsOrderStatus.EOF Then
			
			  set objCmd = Server.CreateObject("ADODB.command")
			  objCmd.ActiveConnection = DataConn
			  objCmd.CommandText = "UPDATE sent_items SET shipped = 'Pending shipment' WHERE ID = ?"
			  objCmd.Parameters.Append(objCmd.CreateParameter("invoiceID",3,1,10,request.querystring("invoice")))
			  objCmd.Execute()
		
		End If

	rsOrderStatus.Close()
	Set rsOrderStatus = Nothing

end if ' if we've received the item

if request.form("backorder") = "yes" then

	  set objCmd = Server.CreateObject("ADODB.command")
	  objCmd.ActiveConnection = DataConn
	  objCmd.CommandText = "UPDATE TBL_OrderSummary SET backorder = 1 WHERE OrderDetailID = ?"
	  objCmd.Parameters.Append(objCmd.CreateParameter("OrderDetailID",3,1,10,request.form("id")))
	  objCmd.Execute()
	  
	  set objCmd = Server.CreateObject("ADODB.command")
	  objCmd.ActiveConnection = DataConn
	  objCmd.CommandText = "SELECT TBL_OrderSummary.OrderDetailID, TBL_OrderSummary.InvoiceID, jewelry.title, sent_items.customer_first, sent_items.email FROM TBL_OrderSummary INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID INNER JOIN sent_items ON TBL_OrderSummary.InvoiceID = sent_items.ID WHERE (TBL_OrderSummary.OrderDetailID = ?)"
	  objCmd.Parameters.Append(objCmd.CreateParameter("OrderDetailID",3,1,10,request.form("id")))
	  Set rsGetItem = objCmd.Execute()
	  
  
mailer_type = request.form("type")
%>
<!--#include virtual="emails/function-send-email.asp"-->
<!--#include virtual="emails/email_variables.asp"-->
<%
end if ' If the item was backordered
%>
