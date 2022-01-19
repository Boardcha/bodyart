<%
if request.cookies("OrderAddonsActive") <> "" then
	var_addons_remove = " AND cart_addon_item = 1"
end if

	'Remove all items from cart after successful payment
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	' cart_save_for_later = 0, initial value
	' cart_save_for_later = 1, it is saved for later
	' cart_save_for_later = 2, it is added back to the cart (to track how many people is checking out with saved for later cart items)	
	objCmd.CommandText = "DELETE FROM tbl_carts WHERE (tbl_carts." & var_db_field & " = ?) " & var_addons_remove & " AND cart_save_for_later <> 1"
	objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10,var_cart_userid))
	objCmd.Execute()

%>