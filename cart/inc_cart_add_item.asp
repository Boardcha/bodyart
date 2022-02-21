<%
' Detect additions from link or form submit (wishlist or product page)

if request.querystring("qty") <> "" then
	var_add_cart_qty = request.querystring("qty")
else
	var_add_cart_qty = request.form("qty")
end if

if request.querystring("DetailID") <> "" then
	var_add_cart_detailId = request.querystring("DetailID")
else
	var_add_cart_detailId = request.form("DetailID")
end if

if request.querystring("anodID") <> "" then
	var_anodID = request.querystring("anodID")
else
	var_anodID = request.form("anodID")
end if

' response.write "preorders " & request.form("preorders") 
if request.form("preorders") <> "" and request.form("preorders") <> "undefined" then
	var_add_cart_preorders = request.form("preorders")
else
	var_add_cart_preorders = ""
end if


' Add-on items feature (adding items for an order that has not shipped out yet)
if request.cookies("OrderAddonsActive") <> "" then
	var_addon = 1
else
	var_addon = 0
end if

'Check to see if it's a gift certificate and format gift certificate data
if request.form("ProductID") = 2424 then
	var_add_cart_qty = 1
	var_add_cart_detailId = request("DetailID")
	var_add_cart_preorders = request.form("email") & "{}" & request.form("your-name") & "{}" & request.form("gift-message") & "{}" & request.form("name")
end if

	'Add item to cart
	if request.form("DetailID") <> "" or request.querystring("DetailID") <> "" then
	
		' Get cart contents to check for duplicate
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT tbl_carts.cart_id, tbl_carts.cart_preorderNotes, tbl_carts.cart_custId, tbl_carts.cart_qty, tbl_carts.cart_save_for_later, ProductDetails.ProductDetailID, jewelry.customorder, jewelry.ProductID FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN tbl_carts ON ProductDetails.ProductDetailID = tbl_carts.cart_detailId WHERE (tbl_carts." & var_db_field & " = ?)"
		objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10,var_cart_userid))
		Set rs_getCart = objCmd.Execute()
		
			duplicate_cartid = 0			
			Do While Not rs_getCart.EOF
				
			' Do not check for duplicate if it's a gift certificate
			if request.form("ProductID") <> 2424 then
				
				if IsNull(rs_getCart.Fields.Item("cart_preorderNotes").Value) then
					duplicate_preorder_notes = ""
				else
					duplicate_preorder_notes = rs_getCart.Fields.Item("cart_preorderNotes").Value
				end if	
					

					' Check for same values in database cart to combine then and avoid duplicates
					if cstr(var_add_cart_detailId) = cstr(rs_getCart.Fields.Item("ProductDetailID").Value) AND cstr(var_add_cart_preorders) = cstr(duplicate_preorder_notes) then
						duplicate_cartid = rs_getCart.Fields.Item("cart_id").Value
						
						var_saved_status = 0
						' Is item a "save for later" item
						var_saved_status = rs_getCart.Fields.Item("cart_save_for_later").Value
					end if
					
			end if ' Do not check for duplicate if it's a gift certificate
			rs_getCart.MoveNext()
			Loop
			
		if duplicate_cartid = 0 then ' only add a new item to DB if no duplicate is found
		
			var_wishlistID = 0
			if request.querystring("ID") <> "" then
				var_wishlistID = request.querystring("ID")
			end if
			if request.form("ID") <> "" then
				var_wishlistID = request.form("ID")
			end if
		
			set objCmd = Server.CreateObject("ADODB.command")
			objCmd.ActiveConnection = DataConn
			objCmd.CommandText = "INSERT INTO tbl_carts (cart_qty, cart_detailID, " & var_db_field & ", cart_preorderNotes, cart_dateAdded, cart_wishlistid, cart_ip_country, cart_addon_item, anodID) VALUES (?,?,?,?,?,?,?,?,?)"
			objCmd.Parameters.Append(objCmd.CreateParameter("cart_qty",3,1,10,var_add_cart_qty))
			objCmd.Parameters.Append(objCmd.CreateParameter("detailID",3,1,10,var_add_cart_detailId))
			objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10,var_cart_userid))
			objCmd.Parameters.Append(objCmd.CreateParameter("cart_preorderNotes",200,1,2000,var_add_cart_preorders))
			objCmd.Parameters.Append(objCmd.CreateParameter("cart_dateAdded",200,1,30,now()))
			objCmd.Parameters.Append(objCmd.CreateParameter("wishlistID",3,1,10,var_wishlistID))
			objCmd.Parameters.Append(objCmd.CreateParameter("ip_country",200,1,5,Request.ServerVariables("HTTP_NGX_GEOIP2_COUNTRYCODE")))
			objCmd.Parameters.Append(objCmd.CreateParameter("cart_addon_item",3,1,2,var_addon))
			objCmd.Parameters.Append(objCmd.CreateParameter("anodID",3,1,10,var_anodID))
			objCmd.Execute()
		
		else ' if a duplicate is found, then update the qty in the current row
		
			' If it's a save for later item that the customer forgot about, then just move it back into the cart and do not add the qty to it either
			if var_saved_status = 1 then
			
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				' cart_save_for_later = 0, initial value
				' cart_save_for_later = 1, it is saved for later
				' cart_save_for_later = 2, it is added back to the cart (to track how many people is checking out with saved for later cart items)
				objCmd.CommandText = "UPDATE tbl_carts SET cart_save_for_later = 2 WHERE " & var_db_field & " = ? AND cart_id = ?"
					
				objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10,var_cart_userid))
				objCmd.Parameters.Append(objCmd.CreateParameter("cart_id",3,1,10,duplicate_cartid))
				objCmd.Execute()
			
			else ' add qty to duplicate
			
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE tbl_carts SET cart_qty = cart_qty + ? WHERE " & var_db_field & " = ? AND cart_id = ?"
					
				objCmd.Parameters.Append(objCmd.CreateParameter("cart_qty",3,1,10,var_add_cart_qty))
				objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10,var_cart_userid))
				objCmd.Parameters.Append(objCmd.CreateParameter("cart_id",3,1,10,duplicate_cartid))
				objCmd.Execute()
			
			end if ' duplicate found
			
		end if ' only add a new item to DB if no duplicate is found 
		
	end if ' end add item to cart
%>