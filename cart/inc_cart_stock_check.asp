<%
	stock_changes = ""
	cart_count = 0
	session("stock_display") = ""
	
	
	Do While Not rs_getCart.EOF
		
		'Assign variables depending on whether an item has been changed or not. We only want to check the stock on the most recently item where the qty was changed in the cart
		if (request.querystring("qty") <> "") and (CLng(request.querystring("detailid")) = CLng(rs_getCart.Fields.Item("cart_detailID").Value)) then
		'	response.write "Updated from qty change"
			var_cart_qty = request.querystring("qty") ' & "-URL"
		else
		'	response.write "Regular stock check"
			var_cart_qty = rs_getCart.Fields.Item("cart_qty").Value ' & "-DB"
		end if
		
		' TO AVOID customers buying products out from under customers that have already added the same item to their cart 
		' We need to calculate quantity by subtracting the items that have been added to the cart in last 15mins from the actual stock count
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT COALESCE(SUM(cart_qty), 0) as cart_qty FROM tbl_carts WHERE cart_detailID = ? AND cart_dateAdded > DATEADD(mi, -15, GETDATE())"
		objCmd.Parameters.Append(objCmd.CreateParameter("cart_detailID",3,1,10, rs_getCart.Fields.Item("cart_detailID").Value))
		Set rs_getTotalQtyinCart = objCmd.Execute()
		If Not rs_getTotalQtyinCart.EOF Then
			'Get total item count for this product variation based on product detail id, EXCLUDING customer's own items
			item_count_in_cart = CLng(rs_getTotalQtyinCart("cart_qty").value) - CLng(var_cart_qty)
			if item_count_in_cart < 0 Then item_count_in_cart = 0
		Else
			item_count_in_cart = 0 
		End If
		
		dynamic_stock_quantity = CLng(rs_getCart("qty").Value) - item_count_in_cart
		If dynamic_stock_quantity < 0 Then dynamic_stock_quantity = 0
		
		'Response.Write "SELECT COALESCE(SUM(cart_qty), 0) as cart_qty FROM tbl_carts WHERE cart_detailID = ? AND cart_dateAdded > DATEADD(mi, -15, GETDATE()) AND cart_custId !=" & var_cart_userid
		'Response.Write "<br>
		'Response.Write var_cart_userid & " - " & CLng(var_cart_qty) & " - " & dynamic_stock_quantity & " - " & CLng(rs_getCart("qty").Value) & " - " & item_count_in_cart
		'Response.End
		
		if rs_getCart.Fields.Item("cart_qty").Value <= 0 then
			
		' Extra safeguard to make sure people can't have negative qty's of items in their cart (even if we have the item in stock)
		'	response.write "Negative qty in cart, set to 0"
			var_qty_update_value = 0
			
		elseif (CLng(var_cart_qty) > dynamic_stock_quantity) then
			
			'If requesting more than what we have on hand, then set to what we have in stock
			var_qty_update_value = dynamic_stock_quantity
			
		end if

		' bug testing
		stock_qtys = stock_qtys & "Querystring ID: " & request.querystring("detailid") & "   Cart id: " & rs_getCart.Fields.Item("cart_detailID").Value & "   Requested qty: " & var_cart_qty & "   My cart qty: " &  rs_getCart.Fields.Item("cart_qty").Value & "  QTY: " &  rs_getCart.Fields.Item("qty").Value & "<br/>"
		' response.write stock_qtys
		
		'Check stock on items
		if (CLng(var_cart_qty) > dynamic_stock_quantity) then
		
			if dynamic_stock_quantity <> 0 then
				var_stock_status = " Quantity changed to " & dynamic_stock_quantity
			else ' if the item is out give delete notice
				var_stock_status = " Deleted from cart "
			end if
				' Value to be passed to another page to update qty field on view cart page via ajax
				var_orig_qty = dynamic_stock_quantity
			
			stock_changes = stock_changes & "<img src=https://s3.amazonaws.com/bodyartforms-products/" & rs_getCart.Fields.Item("picture").Value & "  style=""height: 40px; width: 40px;"">" & var_stock_status & " -- " & rs_getCart.Fields.Item("gauge").Value & " " & rs_getCart.Fields.Item("title").Value & "|"
			
			' Update database with new cart qty value
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "UPDATE tbl_carts SET cart_qty = ? WHERE cart_id = ? AND " & var_db_field & " = ?"
			objCmd.Parameters.Append(objCmd.CreateParameter("cart_qty",3,1,10,var_qty_update_value))
			objCmd.Parameters.Append(objCmd.CreateParameter("cart_id",3,1,10,rs_getCart.Fields.Item("cart_id").Value))
			objCmd.Parameters.Append(objCmd.CreateParameter("cust_id",3,1,10,var_cart_userid))
			objCmd.Execute()

			' Delete AFTER update so it doesn't find an empty record
			if dynamic_stock_quantity = 0 then
								
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "DELETE FROM tbl_carts WHERE cart_id = ? AND " & var_db_field & " = ?"
				objCmd.Parameters.Append(objCmd.CreateParameter("cart_id",3,1,10,rs_getCart.Fields.Item("cart_id").Value))
				objCmd.Parameters.Append(objCmd.CreateParameter("cust_id",3,1,10,var_cart_userid))
				objCmd.Execute()
			
			end if
			
			
		end if ' if customer ordering more than we have on hand


		
		
	rs_getCart.MoveNext()
	Loop


	
if stock_changes <> "" then
	stock_array =split(stock_changes,"|")
	stock_display = ""
		For Each strItem In stock_array
			stock_display = stock_display & "<br/>" & strItem
		Next

		session("stock_display") = stock_display
end if

rs_getCart.ReQuery()
%>