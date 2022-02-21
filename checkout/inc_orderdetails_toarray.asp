<%
'================================================================================================
' START store details into a dynamic multidimensional array
'reDim array_details_2(14,0)
'Dim array_add_new : array_add_new = 0 

'Do while NOT rs_getCart.EOF ' if cart is not emptyh

	array_gauge = ""
	if rs_getCart.Fields.Item("Gauge").Value <> "" then
		array_gauge = Server.HTMLEncode(rs_getCart.Fields.Item("Gauge").Value)
	end if
	
	array_length = ""
	if rs_getCart.Fields.Item("Length").Value <> "" then
		array_length = Server.HTMLEncode(rs_getCart.Fields.Item("Length").Value)
	end if
	
	array_detail = ""
	if rs_getCart.Fields.Item("ProductDetail1").Value <> "" then
		array_detail = Server.HTMLEncode(rs_getCart.Fields.Item("ProductDetail1").Value)
	end if
	
	array_add_new = uBound(array_details_2,2) 
	REDIM PRESERVE array_details_2(14,array_add_new+1) 

	array_details_2(0,array_add_new) = rs_getCart.Fields.Item("ProductDetailID").Value
	array_details_2(1,array_add_new) = rs_getCart.Fields.Item("cart_qty").Value
	array_details_2(2,array_add_new) = rs_getCart.Fields.Item("title").Value 
	array_details_2(3,array_add_new) = array_gauge
	array_details_2(4,array_add_new) = FormatNumber(var_itemPrice_USdollars, -1, -2, -2, -2)
	
	var_preorder_text = ""
	if rs_getCart.Fields.Item("cart_preorderNotes").Value <> "" then
		var_preorder_text = replace(rs_getCart.Fields.Item("cart_preorderNotes").Value,"{}", "   ")
	end if
	
	
	array_details_2(5,array_add_new) = var_preorder_text
	array_details_2(6,array_add_new) = rs_getCart.Fields.Item("ProductID").Value
	array_details_2(7,array_add_new) = "" ' item notes
	array_details_2(8,array_add_new) = FormatNumber(rs_getCart.Fields.Item("wlsl_price").Value, -1, -2, -2, -2)
	array_details_2(9,array_add_new)= rs_getCart.Fields.Item("picture").Value
	array_details_2(10,array_add_new) = array_length
	array_details_2(11,array_add_new) = array_detail
	array_details_2(12,array_add_new) = rs_getCart("free") 
	array_details_2(13,array_add_new) = rs_getCart("cart_save_for_later") 
	array_details_2(14,array_add_new) = rs_getCart("anodID") 
	
'	rs_getCart.MoveNext()

'loop ' if cart is not empty


'================================================================================================
' END store details into a dynamic multidimensional array

%>