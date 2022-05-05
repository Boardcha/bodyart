<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<% 
Details = ""


if request.Form("orderdetailid") <> "" then
	sql_item_id = "AND TBL_OrderSummary.OrderDetailID = " & request.Form("orderdetailid")
end if

if request.Form("brand") <> "" then
	sql_brand = "AND jewelry.brandname ='" & replace(request.form("brand"), "_", " ") & "'"
end if

'response.write "Detail " & sql_item_id & "<br>"
'response.write "sql_brand " & sql_brand & "<br>"

	SqlString = "ALTER VIEW vw_barcodes_preorders AS SELECT TOP (100) PERCENT CAST(dbo.ProductDetails.DetailCode AS varchar(15)) + '' + CAST(dbo.ProductDetails.location AS varchar(15)) AS locationBarcode, dbo.sent_items.ID, dbo.sent_items.shipped, dbo.sent_items.customer_first, dbo.sent_items.customer_last, dbo.jewelry.brandname, dbo.TBL_OrderSummary.qty, dbo.jewelry.title, dbo.ProductDetails.Gauge, dbo.ProductDetails.ProductDetailID, dbo.TBL_OrderSummary.PreOrder_Desc, dbo.TBL_OrderSummary.item_ordered, dbo.TBL_OrderSummary.item_received FROM dbo.jewelry INNER JOIN dbo.TBL_OrderSummary ON dbo.jewelry.ProductID = dbo.TBL_OrderSummary.ProductID INNER JOIN dbo.sent_items ON dbo.TBL_OrderSummary.InvoiceID = dbo.sent_items.ID INNER JOIN dbo.ProductDetails ON dbo.TBL_OrderSummary.DetailID = dbo.ProductDetails.ProductDetailID WHERE (dbo.TBL_OrderSummary.item_ordered = 1) AND (dbo.TBL_OrderSummary.item_received = 0) AND (dbo.sent_items.shipped = N'ON ORDER') " & sql_item_id & " " & sql_brand & " ORDER BY dbo.sent_items.ID"
	Set rsBarcodes = DataConn.Execute(SqlString)

Set rsBarcodes = Nothing 
DataConn.Close()
%>
