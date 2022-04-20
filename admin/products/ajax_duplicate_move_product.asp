<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'duplicate product
if request.form("productid") <> "" and request.form("duplicate")= "product-only" then
	
		set CopyProduct = Server.CreateObject("ADODB.Command")
		CopyProduct.ActiveConnection = DataConn
		CopyProduct.CommandText = "INSERT INTO jewelry(jewelry, type, title, description, picture, picture_400, largepic, material, blackline, internal, customorder, brandname, retainer, pair, flare_type, active, new_page_date, date_added, added_by) SELECT jewelry, type, title, description, 'nopic.gif', 'nopic.gif', 'nopic.gif', material, blackline, internal, customorder, brandname, retainer, pair, flare_type, " & 0 & ", '" & now() & "', '" & now() & "', '" & user_name & "' FROM jewelry WHERE ProductID =" & request.form("productid") 
		CopyProduct.Execute() 
		
		Set objCmd = Server.CreateObject ("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT TOP 1 ProductID FROM jewelry ORDER BY ProductID DESC" 
		Set rsGetID = objCmd.Execute()
     
%>
{  
   "productid":"<%= rsGetID.Fields.Item("ProductID").Value %>"
}
<%
Set rsGetID = Nothing		
end if ' duplicate product


'duplicate product PLUS all it's details  --------------------------------------
if request.form("productid") <> "" and request.form("duplicate")= "all"  then
	
		set CopyProduct = Server.CreateObject("ADODB.Command")
		CopyProduct.ActiveConnection = DataConn
		CopyProduct.CommandText = "INSERT INTO jewelry(jewelry, type, title, description, picture, picture_400, largepic, material, blackline, internal, customorder, brandname, retainer, pair, flare_type, active, new_page_date, date_added, added_by) SELECT jewelry, type, title, description, 'nopic.gif', 'nopic.gif', 'nopic.gif', material, blackline, internal, customorder, brandname, retainer, pair, flare_type, " & 0 & ", '" & now() & "', '" & now() & "', '" & user_name & "' FROM jewelry WHERE ProductID =" & request.form("productid") 
		CopyProduct.Execute() 
		
		Set objCmd = Server.CreateObject ("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT TOP 1 ProductID FROM jewelry ORDER BY ProductID DESC" 
		Set rsGetID = objCmd.Execute()

'response.write "Form product id: " & request.form("productid")
'response.write "Database product id: " & rsGetID.Fields.Item("ProductID").Value
	
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO ProductDetails(active, ProductID, price, wlsl_price, ProductDetail1, detail_code, stock_qty, reorder_amt, item_order, price_wholesale, ShippingRestriction, DetailCode, Gauge, Length, Color, DateAdded, qty, colors, detail_materials, wearable_material) SELECT active, " & rsGetID.Fields.Item("ProductID").Value & ", price, wlsl_price, ProductDetail1, detail_code, stock_qty, reorder_amt, item_order, price_wholesale, ShippingRestriction, 0, Gauge, Length, Color, '" & now() & "', 0, colors, detail_materials, wearable_material FROM ProductDetails WHERE ProductID = " & request.form("productid") 
		objCmd.Execute()
		
		'Update location to newly created detail ID'set				
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE ProductDetails SET location = ProductDetailID WHERE ProductID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,10, rsGetID.Fields.Item("ProductID").Value))
		objCmd.Execute()	
     
%>
{  
   "productid":"<%= rsGetID.Fields.Item("ProductID").Value %>"
}
<%
Set rsGetID = Nothing		
end if ' duplicate product PLUS all it's details --------------------------------


'move detail row to a new product
if request.form("toggle_type") = "move" then

		detail_array =split(request.form("details"),",")
		For Each strItem In detail_array

			if strItem <> "" then 
				set objCmd = Server.CreateObject("ADODB.Command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE ProductDetails SET ProductID = " + request.form("move_to_id") + "  WHERE ProductDetailID = " + strItem
				objCmd.Execute()

				'======= RE-ASSIGN DETAIL IMAGE TO NEW PRODUCT IF AN IMAGE WAS ASSIGNED ========
				set objCmd = Server.CreateObject("ADODB.Command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE tbl_images SET product_id = ? WHERE product_id = ?" 
				objCmd.Parameters.Append(objCmd.CreateParameter("movetoId",3,1,10, request.form("move_to_id") ))
				objCmd.Parameters.Append(objCmd.CreateParameter("moveFromId",3,1,10, request.form("orig_productid") ))
				objCmd.Execute()

			end if ' make sure a detail id is provided to write db
			
		Next
%>
{  
   "productid":"<%= request.form("move_to_id") %>"
}
<%	
end if ' move detail to a new product

'copy detail row
if request.form("toggle_type") = "copy" then

'response.write "ID " & request.form("move_to_id") & " colors: " & var_colors &  " materials: " & var_materials & " wearable: " & request.form("wearable")
'response.write request.form("details")
		' break out form variables into details and rebuild WHERE statement
		detail_array =split(request.form("details"),",")
		For Each strItem In detail_array
	
		'response.write strItem + " ID "

			if strItem <> "" then 
				set objCmd = Server.CreateObject("ADODB.Command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "INSERT INTO ProductDetails(active, ProductID, price, wlsl_price, ProductDetail1, qty, detail_code, stock_qty, reorder_amt, free, item_order, price_wholesale, DateRestocked, ShippingRestriction,  Free_QTY, ReOrder_Amount, Gauge, Length, Color, DateAdded, img_id, colors, detail_materials, wearable_material, DetailCode) SELECT active, ?, price, wlsl_price, ProductDetail1, 0, detail_code, stock_qty, reorder_amt, free, item_order, price_wholesale, DateRestocked, ShippingRestriction,  Free_QTY, ReOrder_Amount, Gauge, Length, Color, '" & now() & "', 0, colors, detail_materials, wearable_material, ? FROM ProductDetails WHERE ProductDetailID = " + strItem 
				objCmd.Parameters.Append(objCmd.CreateParameter("movetoId",3,1,10,request.form("move_to_id")))
				objCmd.Parameters.Append(objCmd.CreateParameter("section",3,1,10,request.form("section")))
				objCmd.Execute()
				
				' retrieve newest product detail ID # to copy into location field
				Set objCmd = Server.CreateObject ("ADODB.Command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "SELECT TOP 1 ProductDetailID FROM ProductDetails ORDER BY ProductDetailID DESC" 
				Set rsGetID = objCmd.Execute()
				
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE ProductDetails SET location = ? WHERE ProductDetailID = ?"
				objCmd.Parameters.Append(objCmd.CreateParameter("location",3,1,10, rsGetID.Fields.Item("ProductDetailID").Value))
				objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10, rsGetID.Fields.Item("ProductDetailID").Value))
				objCmd.Execute()	
				
			end if ' make sure a detail id is provided to write db
			
		Next

%>
{  
   "productid":"<%= request.form("move_to_id") %>"
}
<%		
end if ' if copying detail into a new product

DataConn.Close()
%>