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
' START store details into a dynamic multidimensional array
if request.cookies("gaugecard") <> "no" and var_other_items = 1 then ' add gauge card to array -----------------------

		array_add_new = uBound(array_details_2,2)
		REDIM PRESERVE array_details_2(14,array_add_new+1) 

		array_details_2(0,array_add_new) = "5461"
		array_details_2(1,array_add_new) = "1" ' qty
		array_details_2(2,array_add_new) = "Gauge card" 
		array_details_2(3,array_add_new) = "" ' gauge
		array_details_2(4,array_add_new) = "0" ' price
		array_details_2(5,array_add_new) = "" ' custom order specs
		array_details_2(6,array_add_new) = "1430" ' product id
		array_details_2(7,array_add_new) = "FREE" ' item notes
		array_details_2(8,array_add_new) = ".02" ' wholesale price
		array_details_2(9,array_add_new) = "baf-gauge-card-brown-stock-1430.jpg" '===== image info
		array_details_2(10,array_add_new) = "" '===== length
		array_details_2(11,array_add_new) = "" '===== detail
		array_details_2(12,array_add_new) = 30 

end if ' add gauge card to array ---------------------------------------------------------------

if request.cookies("oringsid") <> "" then ' add o-ring item to details array ------------------
	Do While Not rsGetOrings.EOF
	
	if cStr(rsGetOrings.Fields.Item("ProductDetailID").Value) = request.cookies("oringsid") then ' only retrive item customer selected

		array_add_new = uBound(array_details_2,2)
		REDIM PRESERVE array_details_2(14,array_add_new+1) 

		array_details_2(0,array_add_new) = rsGetOrings.Fields.Item("ProductDetailID").Value
		array_details_2(1,array_add_new) = "4"
		array_details_2(2,array_add_new) = rsGetOrings.Fields.Item("title").Value 
		array_details_2(3,array_add_new) = Server.HTMLEncode(rsGetOrings.Fields.Item("Gauge").Value)
		array_details_2(4,array_add_new) = "0"
		array_details_2(5,array_add_new) = "" 'custom order specs
		array_details_2(6,array_add_new) = rsGetOrings.Fields.Item("ProductID").Value
		array_details_2(7,array_add_new) = "FREE" ' item notes
		array_details_2(8,array_add_new) = ".05" ' wholesale price
		array_details_2(9,array_add_new) = rsGetOrings("picture") '===== image info
		array_details_2(10,array_add_new) = "" '===== length
		array_details_2(11,array_add_new) = "" '===== detail
		array_details_2(12,array_add_new) = 30 
		
	end if ' find matching information for stored cookie id

	rsGetOrings.MoveNext()
	Loop
	Set rsGetOrings = Nothing
	
end if ' add o-ring to array -------------------------------------------------------------------

if request.cookies("stickerid") <> "" then ' add sticker to details array ----------------------
	Do While Not rsGetFree.EOF
	
	if cStr(rsGetFree.Fields.Item("ProductDetailID").Value) = request.cookies("stickerid") then ' only retrive item customer selected

		array_add_new = uBound(array_details_2,2)
		REDIM PRESERVE array_details_2(14,array_add_new+1) 

		array_details_2(0,array_add_new) = rsGetFree.Fields.Item("ProductDetailID").Value
		array_details_2(1,array_add_new) = "1"
		array_details_2(2,array_add_new) = rsGetFree.Fields.Item("title").Value 
		array_details_2(3,array_add_new) = ""
		array_details_2(4,array_add_new) = "0"
		array_details_2(5,array_add_new) = "" ' custom order specs
		array_details_2(6,array_add_new) = rsGetFree.Fields.Item("ProductID").Value
		array_details_2(7,array_add_new) = "FREE" ' item notes
		array_details_2(8,array_add_new) = ".04" ' wholesale price
		array_details_2(9,array_add_new) = rsGetFree("picture") '===== image info
		array_details_2(10,array_add_new) = "" '===== length
		array_details_2(11,array_add_new) = rsGetFree.Fields.Item("ProductDetail1").Value
		array_details_2(12,array_add_new) = 30 
		
	end if ' find matching information for stored cookie id

	rsGetFree.MoveNext()
	Loop
	rsGetFree.MoveFirst()
	
end if ' add sticker to array -------------------------------------------------------------------

gift_count = 1

do until gift_count = 7 ' loop through free gifts
	if request.cookies("freegift" & gift_count & "id") <> "" then ' add 1st free gift to details array -----------------
		Do While Not rsGetFree.EOF
		
		if cStr(rsGetFree.Fields.Item("ProductDetailID").Value) = request.cookies("freegift" & gift_count & "id") then ' only retrieve item customer selected
		
				free_price = 0			
			if inStr(rsGetFree.Fields.Item("ProductDetail1").Value, "USE NOW") <= 0 Then

			if rsGetFree.Fields.Item("free").Value <= var_subtotal_after_discounts then ' fraud check
			
				array_add_new = uBound(array_details_2,2)
				REDIM PRESERVE array_details_2(14,array_add_new+1) 

				array_details_2(0,array_add_new) = rsGetFree.Fields.Item("ProductDetailID").Value
				array_details_2(1,array_add_new) = rsGetFree.Fields.Item("Free_QTY").Value
				array_details_2(2,array_add_new) = rsGetFree.Fields.Item("title").Value 
				array_details_2(3,array_add_new) = rsGetFree.Fields.Item("Gauge").Value
				array_details_2(4,array_add_new) = free_price
				array_details_2(5,array_add_new) = "" ' custom order specs
				array_details_2(6,array_add_new) = rsGetFree.Fields.Item("ProductID").Value
				array_details_2(7,array_add_new) = "FREE" ' item notes
				array_details_2(8,array_add_new) = rsGetFree.Fields.Item("wlsl_price").Value ' wholesale price
				array_details_2(9,array_add_new) = rsGetFree("picture") '===== image info
				array_details_2(10,array_add_new) = rsGetFree.Fields.Item("Length").Value
				array_details_2(11,array_add_new) = rsGetFree.Fields.Item("ProductDetail1").Value
				array_details_2(12,array_add_new) = 30 
			
			end if ' fraud check
			end if ' only write non USE NOW credit items
		
		end if ' find matching information for stored cookie id

		rsGetFree.MoveNext()
		Loop
		rsGetFree.MoveFirst()
		
	end if ' END add 1st free gift to array ---------------------------------------------------------

	gift_count = gift_count + 1
loop ' loop through free gifts
Set rsGetFree = nothing


'================================================================================================
' END store details into a dynamic multidimensional array


' only allow free items on regular orders (not orders where customers are adding on items)
end if ' if not OrderAddonsActive
%>