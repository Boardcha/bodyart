<%
Function GetOrderItems(InvoiceID)
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT InvoiceID, ProductID, DetailID, title, ProductDetail1, Gauge, Length, stock_qty, OrderDetailID, email, customer_first, title, qty, ProductDetail1, ProductDetailID, item_price, PreOrder_Desc, picture, free, type, title FROM dbo.QRY_OrderDetails WHERE InvoiceID = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,20, InvoiceID))
	Set rsGetInfo = objCmd.Execute()

	'================================================================================================
	' START store details into a dynamic multidimensional array
	reDim array_details_2(12,0)

    while NOT rsGetInfo.EOF

		array_gauge = ""
		if rsGetInfo("Gauge") <> "" then
			array_gauge = Server.HTMLEncode(rsGetInfo("Gauge"))
		end if
		
		array_length = ""
		if rsGetInfo("Length") <> "" then
			array_length = Server.HTMLEncode(rsGetInfo("Length"))
		end if
		
		array_detail = ""
		if rsGetInfo("ProductDetail1") <> "" then
			array_detail = Server.HTMLEncode(rsGetInfo("ProductDetail1"))
		end if
		
		array_add_new = uBound(array_details_2,2) 
		REDIM PRESERVE array_details_2(12,array_add_new+1) 

		array_details_2(0,array_add_new) = rsGetInfo("ProductDetailID")
		array_details_2(1,array_add_new) = rsGetInfo("qty")
		array_details_2(2,array_add_new) = rsGetInfo("title") 
		array_details_2(3,array_add_new) = array_gauge
		array_details_2(4,array_add_new) = FormatNumber(rsGetInfo("item_price"), -1, -2, -2, -2)
		
		var_preorder_text = ""
		if rsGetInfo("PreOrder_Desc") <> "" then
			var_preorder_text = replace(rsGetInfo("PreOrder_Desc"),"{}", "   ")
		end if
		
		array_details_2(5,array_add_new) = var_preorder_text
		array_details_2(6,array_add_new) = rsGetInfo("ProductID")
		array_details_2(7,array_add_new) = "" ' item notes
		array_details_2(8,array_add_new) = "" '=== anodization fee
		array_details_2(9,array_add_new)= rsGetInfo("picture")
		array_details_2(10,array_add_new) = array_length
		array_details_2(11,array_add_new) = array_detail
		array_details_2(12,array_add_new) = rsGetInfo("free") 

    rsGetInfo.MoveNext()
    Wend

    GetOrderItems = array_details_2
		
	'================================================================================================
	' END store details into a dynamic multidimensional array

End Function
%>