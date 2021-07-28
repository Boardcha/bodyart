<%
' Check to see if the item has a wishlist ID, and if it does update the wishlist table to purchased = 1
if rs_getCart.Fields.Item("cart_wishlistid").Value <> 0 then
		
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE wishlist SET purchased = 1 WHERE ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("wishlistID",3,1,12,rs_getCart.Fields.Item("cart_wishlistid").Value))
	objCmd.Execute()
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE TBL_OrderSummary SET WishlistID = ? WHERE InvoiceID = ? AND DetailID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("wishlistID",3,1,12,rs_getCart.Fields.Item("cart_wishlistid").Value))
	objCmd.Parameters.Append(objCmd.CreateParameter("invoice_id",3,1,15,session("invoiceid")))
	objCmd.Parameters.Append(objCmd.CreateParameter("cart_detailId",3,1,12,rs_getCart.Fields.Item("cart_detailId").Value))
	objCmd.Execute()
	
end if ' if item has a wishlist ID associated with it
%>