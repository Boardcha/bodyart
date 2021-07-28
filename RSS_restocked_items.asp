<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/sql_connection.asp" -->
<%

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TOP 100 jewelry.ProductID, jewelry.title, jewelry.picture, ProductDetails.ProductDetail1, ProductDetails.price, ProductDetails.DateRestocked, ProductDetails.active, jewelry.active AS activeMain, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.qty, jewelry.SaleDiscount, jewelry.brandname, jewelry.type, jewelry.date_added, jewelry.material, jewelry.internal, jewelry.customorder, jewelry.jewelry, jewelry.flare_type FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID WHERE (ProductDetails.active = 1) AND (jewelry.active = 1) AND DateRestocked > (getdate() - 45) ORDER BY DateRestocked DESC"
'objCmd.Parameters.Append(objCmd.CreateParameter("Dates",200,1,12,Date()))
Set rsGetItems = objCmd.Execute()
rsGetItems_numRows = 0
%>
<%
'***********************
' http://www.dwzone.it
' Rss Writer
' Version 1.1.4
' Start Code
'***********************
Dim dwzRss_rsGetItems
Set dwzRss_rsGetItems = new dwzRssExport
dwzRss_rsGetItems.Init
dwzRss_rsGetItems.SetEncoding "utf-8;65001;Unicode (UTF-8)"
dwzRss_rsGetItems.SetPubDate "true"
dwzRss_rsGetItems.SetTitle "BAF restocked items"
dwzRss_rsGetItems.SetDescription "BAF restocked items"
dwzRss_rsGetItems.SetLink ""
dwzRss_rsGetItems.SetRecordset rsGetItems
dwzRss_rsGetItems.SetItemTitle "title"
dwzRss_rsGetItems.SetItemDescription "ProductDetail1"
dwzRss_rsGetItems.SetItemLink "ProductID"
dwzRss_rsGetItems.SetItemLinkText "http://www.bodyartforms.com/productdetails.asp?ProductID="
dwzRss_rsGetItems.SetItemAuthor ""
dwzRss_rsGetItems.SetItemPubDate ""
dwzRss_rsGetItems.SetFileName ""
dwzRss_rsGetItems.SetNumberOfRecord "ALL"
dwzRss_rsGetItems.SetStartOn "ONLOAD", ""
dwzRss_rsGetItems.SetTimeZone "-0600"
dwzRss_rsGetItems.SetFeedImage "__@_@____@_@__"
dwzRss_rsGetItems.SetAdditionalInfo "en__@_@__RSS__@_@__1440__@_@____@_@____@_@____@_@__"
dwzRss_rsGetItems.Execute()
'***********************
' http://www.dwzone.it
' Rss Writer
' End Code
'***********************
%>
<html>
<head>
<title>Bodyartforms recently restocked items [RSS feed]</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body>
</body>
</html>
<%
rsGetItems.Close()
%>
<!--#include file="dwzExport/RssExport.asp" -->
