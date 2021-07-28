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
				DetailAdd = "dbo.sent_items.ID = " & DetailArray(i)
			else
				DetailAdd = " OR dbo.sent_items.ID = " & DetailArray(i)
			end if
			DetailsBuild = DetailsBuild & DetailAdd
		Next
			Details = " AND (" & DetailsBuild  &  ")"
	end if

	if request.Form("DetailSort") = "Greater" then
		Details = "AND (dbo.sent_items.ID > " & request.Form("Details") & ")"
	end if
	
	' greater than and less than
		if request.Form("DetailSort") = "GreaterLess" then
		Details = "AND (dbo.sent_items.ID >= " & request.Form("Details") & ") AND (dbo.sent_items.ID <= " & request.Form("Details2") & ")"
	end if

else
Details = ""
end if

if request.Form("brand") <> "" then

SqlString = "ALTER VIEW vw_barcodes_preorders AS SELECT TOP (100) PERCENT CAST(dbo.ProductDetails.DetailCode AS varchar(15)) + '' + CAST(dbo.ProductDetails.location AS varchar(15)) AS locationBarcode, dbo.sent_items.ID, dbo.sent_items.shipped, dbo.sent_items.customer_first, dbo.sent_items.customer_last, dbo.jewelry.brandname, dbo.TBL_OrderSummary.qty, dbo.jewelry.title, dbo.ProductDetails.Gauge, dbo.ProductDetails.ProductDetailID, dbo.TBL_OrderSummary.PreOrder_Desc, dbo.TBL_OrderSummary.item_ordered, dbo.TBL_OrderSummary.item_received FROM dbo.jewelry INNER JOIN dbo.TBL_OrderSummary ON dbo.jewelry.ProductID = dbo.TBL_OrderSummary.ProductID INNER JOIN dbo.sent_items ON dbo.TBL_OrderSummary.InvoiceID = dbo.sent_items.ID INNER JOIN dbo.ProductDetails ON dbo.TBL_OrderSummary.DetailID = dbo.ProductDetails.ProductDetailID WHERE (dbo.TBL_OrderSummary.item_ordered = 1) AND (dbo.TBL_OrderSummary.item_received = 0) AND (dbo.sent_items.shipped = N'ON ORDER') " & Details & " AND (dbo.jewelry.brandname ='" & request.form("brand") & "') ORDER BY dbo.sent_items.ID"
Set rsBarcodes = DataConn.Execute(SqlString)

end if 
%>
<html>
<head>

<title>Update barcodes</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<div class="alert alert-success">
Query updated ... you can now print barcodes with the NiceLabel program
</div>
</div>
</body>
</html>
<% Set rsBarcodes = Nothing 
DataConn.Close()
%>
