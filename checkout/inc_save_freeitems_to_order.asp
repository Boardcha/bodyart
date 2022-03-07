<%
' only allow free items on regular orders (not orders where customers are adding on items)
if request.cookies("OrderAddonsActive") = "" then

' ------- Get FREE items
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT jewelry.title, jewelry.picture, ProductDetails.ProductDetail1, ProductDetails.qty, ProductDetails.free, jewelry.ProductID, ProductDetails.ProductDetailID, ProductDetails.Free_QTY,  ProductDetails.price, ProductDetails.wlsl_price, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.ProductDetail1 FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID WHERE (ProductDetails.qty > 0) AND (ProductDetails.free <> 0) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.active = 1) AND (jewelry.active = 1)"
		Set rsGetFree = objCmd.Execute()
		
' ------- End getting free items


'================================================================================================
' START store details
if request.cookies("gaugecard") <> "no" and var_other_items = 1 then ' add gauge card to array -----------------------

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, ProductID, DetailID, qty, item_price, notes,  item_wlsl_price) VALUES (?,?,?,?,?,?,?)"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,session("invoiceid")))
		objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15, 1430 ))
		objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15, 5461 ))
		objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10, 1 ))
		objCmd.Parameters.Append(objCmd.CreateParameter("item_price",6,1,10, 0 ))
		objCmd.Parameters.Append(objCmd.CreateParameter("item_notes",200,1,50, "FREE" ))
		objCmd.Parameters.Append(objCmd.CreateParameter("item_wlsl_price",6,1,10, ".02" ))
		objCmd.Execute()

end if ' add gauge card to array ---------------------------------------------------------------

if request.cookies("oringsid") <> "" then ' add o-ring item to details array ------------------
	While Not rsGetOrings.EOF
	
	if cStr(rsGetOrings.Fields.Item("ProductDetailID").Value) = request.cookies("oringsid") then ' only retrive item customer selected

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, ProductID, DetailID, qty, item_price, notes,  item_wlsl_price) VALUES (?,?,?,?,?,?,?)"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,session("invoiceid")))
		objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15, rsGetOrings("ProductID") ))
		objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15, rsGetOrings("ProductDetailID") ))
		objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10, 4 ))
		objCmd.Parameters.Append(objCmd.CreateParameter("item_price",6,1,10, 0 ))
		objCmd.Parameters.Append(objCmd.CreateParameter("item_notes",200,1,50, "FREE" ))
		objCmd.Parameters.Append(objCmd.CreateParameter("item_wlsl_price",6,1,10, rsGetOrings("wlsl_price") ))
		objCmd.Execute()
		
	end if ' find matching information for stored cookie id

	rsGetOrings.MoveNext()
	Wend
	
end if ' add o-ring to array -------------------------------------------------------------------

if request.cookies("stickerid") <> "" then ' add sticker to details array ----------------------
	While Not rsGetFree.EOF
	
	if cStr(rsGetFree.Fields.Item("ProductDetailID").Value) = request.cookies("stickerid") then ' only retrive item customer selected

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, ProductID, DetailID, qty, item_price, notes,  item_wlsl_price) VALUES (?,?,?,?,?,?,?)"
		objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,session("invoiceid")))
		objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15, rsGetFree("ProductID") ))
		objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15, rsGetFree("ProductDetailID") ))
		objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10, 1 ))
		objCmd.Parameters.Append(objCmd.CreateParameter("item_price",6,1,10, 0 ))
		objCmd.Parameters.Append(objCmd.CreateParameter("item_notes",200,1,50, "FREE" ))
		objCmd.Parameters.Append(objCmd.CreateParameter("item_wlsl_price",6,1,10, rsGetFree("wlsl_price") ))
		objCmd.Execute()
		
	end if ' find matching information for stored cookie id

	rsGetFree.MoveNext()
	Wend
	rsGetFree.MoveFirst()
	
end if ' add sticker to array -------------------------------------------------------------------

gift_count = 1

do until gift_count = 7 ' loop through free gifts
	if request.cookies("freegift" & gift_count & "id") <> "" then ' add 1st free gift to details array -----------------
		While Not rsGetFree.EOF
		
		if cStr(rsGetFree.Fields.Item("ProductDetailID").Value) = request.cookies("freegift" & gift_count & "id") then ' only retrieve item customer selected
		
				free_price = 0			
			if inStr(rsGetFree.Fields.Item("ProductDetail1").Value, "USE NOW") <= 0 Then

			if rsGetFree.Fields.Item("free").Value <= var_subtotal_after_discounts then ' fraud check
						
				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "INSERT INTO TBL_OrderSummary (InvoiceID, ProductID, DetailID, qty, item_price, notes,  item_wlsl_price) VALUES (?,?,?,?,?,?,?)"
				objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,session("invoiceid")))
				objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15, rsGetFree("ProductID") ))
				objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,15, rsGetFree("ProductDetailID") ))
				objCmd.Parameters.Append(objCmd.CreateParameter("qty",3,1,10, rsGetFree("Free_QTY") ))
				objCmd.Parameters.Append(objCmd.CreateParameter("item_price",6,1,10, 0 ))
				objCmd.Parameters.Append(objCmd.CreateParameter("item_notes",200,1,50, "FREE" ))
				objCmd.Parameters.Append(objCmd.CreateParameter("item_wlsl_price",6,1,10, rsGetFree("wlsl_price") ))
				objCmd.Execute()
			
			end if ' fraud check
			end if ' only write non USE NOW credit items
		
		end if ' find matching information for stored cookie id

		rsGetFree.MoveNext()
		Wend
		rsGetFree.MoveFirst()
		
	end if ' END add 1st free gift to array ---------------------------------------------------------

	gift_count = gift_count + 1
loop ' loop through free gifts
Set rsGetFree = nothing


'================================================================================================
' END store details


' only allow free items on regular orders (not orders where customers are adding on items)
end if ' if not OrderAddonsActive
%>