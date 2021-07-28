<%
if session("guestID") <> "" and session("guestID") <> 0 then

' If any items are in an un-registered shopping cart, then add them into the signed in account
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE tbl_carts SET cart_custID = ?, cart_guest_userid = 0 WHERE cart_guest_userid = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("cust_id",3,1,10,var_our_custid))
	objCmd.Parameters.Append(objCmd.CreateParameter("guest_id",3,1,10,session("guestID")))
	objCmd.Execute()
	
	' Set session variable to notify customer on main account page that items have been moved to account cart
	session("cart_items_transferred") = "yes"
	
end if
%>