<%
' If customer is NOT registered --------------------------------
if Request.Cookies("ID") <> "" then 

	' decrypt customer ID cookie
	Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")

	password = "3uBRUbrat77V"
	data = request.Cookies("ID")

	If len(data) > 5 then ' if
		decrypted = objCrypt.Decrypt(password, data)
	end if

	  if data <> decrypted then
		  CustID_Cookie = decrypted
	  else
		  CustID_Cookie = 0
	  end if

	Set objCrypt = Nothing
	
end if

	'Check stock before updating
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT TOP(1) qty, ProductDetailID FROM ProductDetails WHERE ProductDetailID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("DetailID",3,1,10, request.querystring("detailid")))
		Set rsGetStockAmount = objCmd.Execute()

	'Update qty on item using JAVASCRIPT
	if request.querystring("update") <> "" AND (Cint(request.querystring("qty")) <= Cint(rsGetStockAmount.Fields.Item("qty").Value)) then

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE tbl_carts SET cart_qty = ? WHERE cart_id = ? AND " & var_db_field & " = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("cart_qty",3,1,10, request.querystring("qty")))
		objCmd.Parameters.Append(objCmd.CreateParameter("cart_id",3,1,10, request.querystring("update")))
		objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10, var_cart_userid))
		objCmd.Execute()
		
		update_success = "yes"
		
	else
	
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE tbl_carts SET cart_qty = ? WHERE cart_id = ? AND " & var_db_field & " = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("cart_qty",3,1,10, rsGetStockAmount.Fields.Item("qty").Value))
		objCmd.Parameters.Append(objCmd.CreateParameter("cart_id",3,1,10, request.querystring("update")))
		objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10, var_cart_userid))
		objCmd.Execute()
		
		update_success = "yes"

	end if ' end update qty on item
	
	



'-------- Update qty on item WITHOUT JAVASCRIPT TURNED ON -----------------------------
if request.querystring("update-nojs") = "yes" then
		form_string = request.form
		nojs_update_array =split(form_string,"&")

		For Each strItem In nojs_update_array
			if Instr(strItem, "qty_change_id_") > 0 then ' only pick out details to update qty's on
				nojs_split_qty_array =split(strItem,"=") ' split them out by the querystring = sign
				For Each y In nojs_split_qty_array
					var_split_finalbuild = var_split_finalbuild & Replace(y, "qty_change_id_", "") & ","
				Next
			end if	
		Next

		nojs_update_qtys = split(var_split_finalbuild,",")
	'	response.write "Array size: " & UBound(nojs_update_qtys) & "<br/>"

		j_update = 0
		For c = 1 to UBound(nojs_update_qtys) Step 2 'update db/cookie every 2, only pull detailid and qty

			var_detailid_update = nojs_update_qtys(j_update)
		'	response.write "Item: " & var_detailid_update & "<br/>"
			j_update = j_update + 1

			var_qty_update = nojs_update_qtys(j_update)
		'	response.write "Qty: " & var_qty_update & "<br/>"
			j_update = j_update + 1

						
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE tbl_carts SET cart_qty = ? WHERE cart_id = ? AND cart_custID = ?"
				objCmd.Parameters.Append(objCmd.CreateParameter("cart_qty",3,1,10, var_qty_update))
				objCmd.Parameters.Append(objCmd.CreateParameter("cart_id",3,1,10, var_detailid_update))
				objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10, CustID_Cookie))
				objCmd.Execute()
			

		Next
end if ' check for requests.querystring("update-nojs")
%>