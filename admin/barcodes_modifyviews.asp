<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<% 
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

if request.Form("Details") <> "" then

	if request.Form("DetailSort") = "Equal" then
		Details = ""
		DetailArray = Split(request.Form("Details"),",")
		For i=0 to UBound(DetailArray) 'the UBound function returns # in array
			if i = 0 then
				DetailAdd = "ProductDetails.location = " & DetailArray(i)
			else
				DetailAdd = " OR ProductDetails.location = " & DetailArray(i)
			end if
			DetailsBuild = DetailsBuild & DetailAdd
		Next
			Details = " AND (" & DetailsBuild  &  ")"
	end if

	if request.Form("DetailSort") = "Greater" then
		Details = "AND (ProductDetails.location > " & request.Form("Details") & ")"
	end if
	
	' greater than and less than
		if request.Form("DetailSort") = "GreaterLess" then
		Details = "AND (ProductDetails.location >= " & request.Form("Details") & ") AND (ProductDetails.location <= " & request.Form("Details2") & ")"
	end if

else
Details = ""
end if

if request.Form("Products") <> "" then

		Products = ""
		ProductsArray = Split(request.Form("Products"),",")
		For j=0 to UBound(ProductsArray) 'the UBound function returns # in array
			if j = 0 then
				ProductsAdd = "jewelry.ProductID = " & ProductsArray(j)
			else
				ProductsAdd = " OR jewelry.ProductID = " & ProductsArray(j)
			end if
		ProductsBuild = ProductsBuild & ProductsAdd

		Next

		if request.Form("Details") <> "" then ' build as OR statement if details are entered
			Products= " OR (" & ProductsBuild  &  ")"
		else
			Products= " AND (" & ProductsBuild  &  ")"
		end if	
else
Products = ""
end if

if request.form("new") = "yes" then
	new_items = " AND ProductDetails.DateAdded > GETDATE() - 60"
else
	new_items = ""
end if

'===== UPDATES VIEW TO RETRIEVE PURCHASE ORDER DETAILS FOR LABEL PRINTERS IN OFFICE =================
if request.QueryString("type") = "Order" then

SqlString = "ALTER VIEW QRY_Labels_Orders AS SELECT TOP (100) PERCENT ProductDetails.ProductDetailID, ProductDetails.ProductDetail1, ProductDetails.location,  ProductDetails.BinNumber_Detail, tbl_po_details.po_orderid as 'PurchaseOrderID', ProductDetails.POAmount, ProductDetails.Gauge, ProductDetails.Length, ISNULL(ProductDetails.Gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' + ISNULL(ProductDetails.ProductDetail1,'') + ' ' + jewelry.title as title, jewelry.title as title_sort,  TBL_Barcodes_SortOrder.ID_Description FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE (tbl_po_details.po_orderid = " & request.QueryString("ID") & ") ORDER BY title_sort" 
Set rsBarcodes = DataConn.Execute(SqlString)

end if 

'===== UPDATES VIEW TO RETRIEVE PURCHASE ORDER DETAILS FOR LABEL PRINTERS IN OFFICE =================
if request.QueryString("type") = "new_po_system" then

SqlString = "ALTER VIEW QRY_Labels_Orders AS SELECT TOP (100) PERCENT ProductDetails.ProductDetailID, ProductDetails.ProductDetail1, ProductDetails.location,  ProductDetails.BinNumber_Detail, tbl_po_details.po_orderid as 'PurchaseOrderID', tbl_po_details.po_qty, ProductDetails.Gauge, ProductDetails.Length, ISNULL(ProductDetails.Gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' + ISNULL(ProductDetails.ProductDetail1,'') + ' ' + jewelry.title as title, jewelry.title as title_sort,  TBL_Barcodes_SortOrder.ID_Description FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number  INNER JOIN tbl_po_details ON ProductDetails.ProductDetailID = tbl_po_details.po_detailid WHERE (tbl_po_details.po_orderid = " & request.QueryString("ID") & ") ORDER BY CASE WHEN jewelry LIKE '%plugs%' THEN 1 WHEN jewelry LIKE '%clicker%' THEN 5 WHEN jewelry LIKE '%captive%' THEN 6 WHEN jewelry LIKE '%septum%' THEN 7 WHEN jewelry LIKE '%balls%' THEN 9 WHEN jewelry LIKE '%labret%' THEN 10 WHEN jewelry LIKE '%circular%' THEN 11 WHEN jewelry LIKE '%nose%' THEN 12 WHEN jewelry LIKE '%belly%' THEN 13 WHEN jewelry LIKE '%nipple%' THEN 14 WHEN jewelry LIKE '%barbell%' THEN 15 WHEN jewelry LIKE '%tapers%' THEN 19 WHEN jewelry LIKE '%saddle%' THEN 20 WHEN jewelry LIKE '%earring%' THEN 25 WHEN jewelry LIKE '%twists%' THEN 26 WHEN jewelry LIKE '%curved%' THEN 27 WHEN jewelry LIKE '%weight%' THEN 40 WHEN jewelry LIKE '%hanging%' THEN 45 WHEN jewelry LIKE '%cuff%' THEN 80 ELSE 0 END ASC, title_sort asc" 
Set rsBarcodes = DataConn.Execute(SqlString)

end if 

'===== UPDATES VIEW TO RETRIEVE ALL ITEMS BY PRODUCT ID FOR LABEL PRINTERS IN OFFICE =================
if request.queryString("type") = "new_item_labels" then

	SqlString = "ALTER VIEW QRY_Labels_Orders AS SELECT TOP (100) PERCENT ProductDetails.ProductDetailID, ProductDetails.ProductDetail1, ProductDetails.location,  ProductDetails.BinNumber_Detail, 0 as 'PurchaseOrderID', ProductDetails.qty as 'po_qty', ProductDetails.Gauge, ProductDetails.Length, ISNULL(ProductDetails.Gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' + ISNULL(ProductDetails.ProductDetail1,'') + ' ' + jewelry.title as title, jewelry.title as title_sort,  TBL_Barcodes_SortOrder.ID_Description FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE ProductDetails.ProductID = " & request.form("productid") & " ORDER BY title_sort" 
	Set rsBarcodes = DataConn.Execute(SqlString)

end if 

'=========  LABELS BY PRODUCT DETAIL ID ONLY ============================
if request.queryString("type") = "labels_by_detailid" then

	SqlString = "ALTER VIEW QRY_Barcodes_Regular AS SELECT TOP (100) PERCENT CAST(ProductDetails.DetailCode AS varchar(15)) + '' + CAST(ProductDetails.location AS varchar(15)) AS locationBarcode, ProductDetails.location, ProductDetails.ProductDetailID, ISNULL(ProductDetails.ProductDetail1,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' + jewelry.title AS title, ProductDetails.ProductDetail1,  ProductDetails.Gauge, ProductDetails.Length, TBL_Barcodes_SortOrder.ID_Description FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID LEFT OUTER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE (jewelry.customorder <> 'yes') " & request.form("detailids") & " ORDER BY ProductDetails.ProductDetailID DESC, ProductDetails.location ASC" 
	Set rsBarcodes = DataConn.Execute(SqlString)

end if

'=========  LABELS FOR CUSTOM ORDER SHIPMENTS ============================
if request.queryString("type") = "custom_orders" then

	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "ALTER VIEW vw_barcodes_preorders AS SELECT TOP (100) PERCENT sent_items.ID, ProductDetails.ProductDetailID, ProductDetails.Gauge, ProductDetails.Length, ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '') + ' ' + jewelry.title AS title, PreOrder_Desc, sent_items.customer_first, TBL_OrderSummary.qty FROM ProductDetails INNER JOIN TBL_OrderSummary ON ProductDetails.ProductDetailID = TBL_OrderSummary.DetailID INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID INNER JOIN sent_items ON TBL_OrderSummary.InvoiceID = sent_items.ID INNER JOIN tbl_po_details ON TBL_OrderSummary.OrderDetailID = tbl_po_details.po_invoice_order_detailid WHERE tbl_po_details.po_orderid = " & request.querystring("ID") & " ORDER BY ID DESC" 

	'objCmd.CommandText = "ALTER VIEW vw_barcodes_preorders AS SELECT TOP (100) PERCENT CAST(dbo.ProductDetails.DetailCode AS varchar(15)) + '' + CAST(dbo.ProductDetails.location AS varchar(15)) AS locationBarcode, dbo.sent_items.ID, dbo.sent_items.shipped, dbo.sent_items.customer_first, dbo.sent_items.customer_last, dbo.jewelry.brandname, dbo.TBL_OrderSummary.qty, dbo.jewelry.title, dbo.ProductDetails.Gauge, dbo.ProductDetails.ProductDetailID, dbo.TBL_OrderSummary.PreOrder_Desc, dbo.TBL_OrderSummary.item_ordered, dbo.TBL_OrderSummary.item_received FROM dbo.jewelry INNER JOIN dbo.TBL_OrderSummary ON dbo.jewelry.ProductID = dbo.TBL_OrderSummary.ProductID INNER JOIN dbo.sent_items ON dbo.TBL_OrderSummary.InvoiceID = dbo.sent_items.ID INNER JOIN dbo.ProductDetails ON dbo.TBL_OrderSummary.DetailID = dbo.ProductDetails.ProductDetailID WHERE TBL_OrderSummary.item_ordered = 1 AND TBL_OrderSummary.item_received = 0 AND tbl_po_details.po_orderid = " & request.querystring("ID") & " ORDER BY dbo.sent_items.ID"

	objCmd.Execute 

end if

'======== DEFAULT SETTING
if request.Form("type") <> "" then
	
	If request.Form("type") <> "-" then
		LocationGroup = " AND ProductDetails.DetailCode = " & request.Form("type") & ""
	end if
	If request.Form("type") = "Limited" then
		LocationGroup = " AND (jewelry.type = 'limited' OR jewelry.type = 'Clearance' OR jewelry.type = 'Discontinued' OR jewelry.type = 'One time buy' OR jewelry.type = 'OneDay')"
	end if	
	If request.Form("type") = "-" then
		LocationGroup = ""
	End if

	SqlString = "ALTER VIEW QRY_Barcodes_Regular AS SELECT TOP (100) PERCENT CAST(ProductDetails.DetailCode AS varchar(15)) + '' + CAST(ProductDetails.location AS varchar(15)) AS locationBarcode, ProductDetails.location, ProductDetails.ProductDetailID, ProductDetails.ProductDetail1 + ' ' + ProductDetails.Length + ' ' + jewelry.title AS title, ProductDetails.ProductDetail1,  ProductDetails.Gauge, ProductDetails.Length, TBL_Barcodes_SortOrder.ID_Description FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID LEFT OUTER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE  (jewelry.customorder <> 'yes') " & Details & " " & Products & " " & LocationGroup & " " & new_items & " ORDER BY ProductDetails.ProductDetailID DESC, ProductDetails.location ASC" 
	Set rsBarcodes = DataConn.Execute(SqlString)

end if 
%>
<html>
<head>
<title>Print barcodes</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">

	<h4 class="alert alert-success">Query updated ... you can now print barcodes with the NiceLabel program</h4>
</div>
</body>
</html>
<% Set rsBarcodes = Nothing 
DataConn.Close()
%>
