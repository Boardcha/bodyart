<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' If we've completed the anodization
if request.querystring("completed") = "yes" then

	  set objCmd = Server.CreateObject("ADODB.command")
	  objCmd.ActiveConnection = DataConn
	  objCmd.CommandText = "UPDATE TBL_OrderSummary SET anodized_completed = 1, anodized_date = GETDATE() WHERE OrderDetailID = ?"
	  objCmd.Parameters.Append(objCmd.CreateParameter("OrderDetailID",3,1,10,request.querystring("id")))
	  objCmd.Execute()
	  
	  
	  set objCmd = Server.CreateObject("ADODB.command")
	  objCmd.ActiveConnection = DataConn
	  objCmd.CommandText = "SELECT InvoiceID, OrderDetailID FROM TBL_OrderSummary WHERE InvoiceID = ? AND anodized_completed = 0 AND anodization_id_ordered > 0"
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

end if 

%>
