<%
'======= RETRIEVE ITEM DETAILS FROM ORDER TO DEDUCT INVENTORY QUANTITY ==================

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM TBL_OrderSummary WHERE InvoiceID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10, Session("invoiceid")))
set rsDeductItems = objCmd.Execute()

while NOT rsDeductItems.EOF
		
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE ProductDetails SET qty = qty - ?, DateLastPurchased = '" & date() & "' WHERE ProductDetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("QtyPurchased",3,1,10, rsDeductItems("qty") ))		
	objCmd.Parameters.Append(objCmd.CreateParameter("DetailID",3,1,12, rsDeductItems("DetailID") ))		
	objCmd.Execute()
		
rsDeductItems.MoveNext()
Wend
%>