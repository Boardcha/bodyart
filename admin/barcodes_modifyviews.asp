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
				DetailAdd = "dbo.ProductDetails.location = " & DetailArray(i)
			else
				DetailAdd = " OR dbo.ProductDetails.location = " & DetailArray(i)
			end if
			DetailsBuild = DetailsBuild & DetailAdd
		Next
			Details = " AND (" & DetailsBuild  &  ")"
	end if

	if request.Form("DetailSort") = "Greater" then
		Details = "AND (dbo.ProductDetails.location > " & request.Form("Details") & ")"
	end if
	
	' greater than and less than
		if request.Form("DetailSort") = "GreaterLess" then
		Details = "AND (dbo.ProductDetails.location >= " & request.Form("Details") & ") AND (dbo.ProductDetails.location <= " & request.Form("Details2") & ")"
	end if

else
Details = ""
end if

if request.Form("Products") <> "" then

		Products = ""
		ProductsArray = Split(request.Form("Products"),",")
		For j=0 to UBound(ProductsArray) 'the UBound function returns # in array
			if j = 0 then
				ProductsAdd = "dbo.jewelry.ProductID = " & ProductsArray(j)
			else
				ProductsAdd = " OR dbo.jewelry.ProductID = " & ProductsArray(j)
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


if request.QueryString("type") = "Order" then

SqlString = "ALTER VIEW QRY_Labels_Orders AS SELECT TOP (100) PERCENT ProductDetails.ProductDetailID, ProductDetails.ProductDetail1, ProductDetails.location,  ProductDetails.BinNumber_Detail, tbl_po_details.po_orderid as 'PurchaseOrderID', ProductDetails.POAmount, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.Gauge + ' ' + ProductDetails.Length + ' ' + ProductDetails.ProductDetail1 + ' ' + jewelry.title as title, jewelry.title as title_sort,  TBL_Barcodes_SortOrder.ID_Description FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE (tbl_po_details.po_orderid = " & request.QueryString("ID") & ") ORDER BY title_sort" 
Set rsBarcodes = DataConn.Execute(SqlString)

end if 

if request.QueryString("type") = "new_po_system" then

SqlString = "ALTER VIEW QRY_Labels_Orders AS SELECT TOP (100) PERCENT ProductDetails.ProductDetailID, ProductDetails.ProductDetail1, ProductDetails.location,  ProductDetails.BinNumber_Detail, tbl_po_details.po_orderid as 'PurchaseOrderID', tbl_po_details.po_qty, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.Gauge + ' ' + ProductDetails.Length + ' ' + ProductDetails.ProductDetail1 + ' ' + jewelry.title as title, jewelry.title as title_sort,  TBL_Barcodes_SortOrder.ID_Description FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number  INNER JOIN dbo.tbl_po_details ON dbo.ProductDetails.ProductDetailID = dbo.tbl_po_details.po_detailid WHERE (tbl_po_details.po_orderid = " & request.QueryString("ID") & ") ORDER BY title_sort" 
Set rsBarcodes = DataConn.Execute(SqlString)

end if 

if request.queryString("type") = "new_item_labels" then

	SqlString = "ALTER VIEW QRY_Labels_Orders AS SELECT TOP (100) PERCENT ProductDetails.ProductDetailID, ProductDetails.ProductDetail1, ProductDetails.location,  ProductDetails.BinNumber_Detail, 0 as 'PurchaseOrderID', ProductDetails.qty as 'po_qty', ProductDetails.Gauge, ProductDetails.Length, ProductDetails.Gauge + ' ' + ProductDetails.Length + ' ' + ProductDetails.ProductDetail1 + ' ' + jewelry.title as title, jewelry.title as title_sort,  TBL_Barcodes_SortOrder.ID_Description FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE ProductDetails.ProductID = " & request.form("productid") & " ORDER BY title_sort" 
	Set rsBarcodes = DataConn.Execute(SqlString)

end if 

'=========  LABELS BY PRODUCT DETAIL ID ONLY ============================
if request.queryString("type") = "labels_by_detailid" then

	SqlString = "ALTER VIEW QRY_Barcodes_Regular AS SELECT TOP (100) PERCENT CAST(dbo.ProductDetails.DetailCode AS varchar(15)) + '' + CAST(dbo.ProductDetails.location AS varchar(15)) AS locationBarcode, ProductDetails.location, dbo.ProductDetails.ProductDetailID, dbo.ProductDetails.ProductDetail1 + ' ' + dbo.ProductDetails.Length + ' ' + dbo.jewelry.title AS title, dbo.ProductDetails.ProductDetail1,  dbo.ProductDetails.Gauge, dbo.ProductDetails.Length, dbo.TBL_Barcodes_SortOrder.ID_Description FROM dbo.jewelry INNER JOIN dbo.ProductDetails ON dbo.jewelry.ProductID = dbo.ProductDetails.ProductID LEFT OUTER JOIN dbo.TBL_Barcodes_SortOrder ON dbo.ProductDetails.DetailCode = dbo.TBL_Barcodes_SortOrder.ID_Number WHERE (dbo.jewelry.customorder <> 'yes') " & request.form("detailids") & " ORDER BY dbo.ProductDetails.ProductDetailID DESC, dbo.ProductDetails.location ASC" 
	Set rsBarcodes = DataConn.Execute(SqlString)

end if


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

SqlString = "ALTER VIEW QRY_Barcodes_Regular AS SELECT TOP (100) PERCENT CAST(dbo.ProductDetails.DetailCode AS varchar(15)) + '' + CAST(dbo.ProductDetails.location AS varchar(15)) AS locationBarcode, ProductDetails.location, dbo.ProductDetails.ProductDetailID, dbo.ProductDetails.ProductDetail1 + ' ' + dbo.ProductDetails.Length + ' ' + dbo.jewelry.title AS title, dbo.ProductDetails.ProductDetail1,  dbo.ProductDetails.Gauge, dbo.ProductDetails.Length, dbo.TBL_Barcodes_SortOrder.ID_Description FROM dbo.jewelry INNER JOIN dbo.ProductDetails ON dbo.jewelry.ProductID = dbo.ProductDetails.ProductID LEFT OUTER JOIN dbo.TBL_Barcodes_SortOrder ON dbo.ProductDetails.DetailCode = dbo.TBL_Barcodes_SortOrder.ID_Number WHERE  (dbo.jewelry.customorder <> 'yes') " & Details & " " & Products & " " & LocationGroup & " " & new_items & " ORDER BY dbo.ProductDetails.ProductDetailID DESC, dbo.ProductDetails.location ASC" 
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
