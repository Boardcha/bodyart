<%
'======= RETRIEVE CART ITEMS AND CREATE ARRAY TO SEND IN EMAIL ORDER CONFIRMATION ==================

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TBL_OrderSummary.qty, gauge, length, ProductDetail1, title, ProductDetailID, TBL_OrderSummary.ProductID, picture, PreOrder_Desc, item_price, free, anodization_fee  FROM TBL_OrderSummary INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID WHERE TBL_OrderSummary.InvoiceID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10, Session("invoiceid")))
set rsBuildEmailArray = objCmd.Execute()

reDim array_details_2(11,0)

while NOT rsBuildEmailArray.EOF
		
        email_gauge = ""
        if rsBuildEmailArray("Gauge") <> "" then
        email_gauge = Server.HTMLEncode(rsBuildEmailArray("Gauge"))
        end if

        email_length = ""
        if rsBuildEmailArray("Length") <> "" then
            email_length = Server.HTMLEncode(rsBuildEmailArray("Length"))
        end if

        email_detail = ""
        if rsBuildEmailArray("ProductDetail1") <> "" then
            email_detail = Server.HTMLEncode(rsBuildEmailArray("ProductDetail1"))
        end if

        array_add_new = uBound(array_details_2,2) 
        REDIM PRESERVE array_details_2(11,array_add_new+1) 

        array_details_2(0,array_add_new) = rsBuildEmailArray("ProductDetailID")
        array_details_2(1,array_add_new) = rsBuildEmailArray("qty")
        array_details_2(2,array_add_new) = rsBuildEmailArray("title") 
        array_details_2(3,array_add_new) = email_gauge
        array_details_2(4,array_add_new) = FormatNumber(rsBuildEmailArray("item_price"), -1, -2, -2, -2)

        email_preorder_specs = ""
        if rsBuildEmailArray("PreOrder_Desc") <> "" then
            email_preorder_specs = replace(rsBuildEmailArray("PreOrder_Desc"),"{}", "   ")
        end if

        array_details_2(5,array_add_new) = email_preorder_specs
        array_details_2(6,array_add_new) = rsBuildEmailArray("ProductID")
        array_details_2(7,array_add_new) = "" ' item notes
        array_details_2(8,array_add_new) = rsBuildEmailArray("anodization_fee")
        array_details_2(9,array_add_new)= rsBuildEmailArray("picture")
        array_details_2(10,array_add_new) = email_length
        array_details_2(11,array_add_new) = email_detail
		
rsBuildEmailArray.MoveNext()
Wend
%>