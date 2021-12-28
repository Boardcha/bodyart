<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->

<%
if Request.Form("DetailID") <> "" then ' only process form if something has been scanned in


Set rsGetRestockItems_cmd = Server.CreateObject ("ADODB.Command")
rsGetRestockItems_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRestockItems_cmd.CommandText = "SELECT ProductDetails.ProductDetailID, jewelry.title, ProductDetails.ProductDetail1, ProductDetails.qty, jewelry.ProductID, ProductDetails.Gauge,  ProductDetails.Length, jewelry.picture, tbl_po_details.po_orderid, tbl_po_details.po_qty FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN tbl_po_details ON ProductDetails.ProductDetailID = tbl_po_details.po_detailid WHERE ProductDetailID = ? AND po_orderid = "+Request("po_id")+"" 
rsGetRestockItems_cmd.Prepared = true
rsGetRestockItems_cmd.Parameters.Append rsGetRestockItems_cmd.CreateParameter("param1", 5, 1, -1, Request.Form("DetailID")) ' adDouble

Set rsGetRestockItems = rsGetRestockItems_cmd.Execute

If rsGetRestockItems.Fields.Item("qty").Value <= 0 then
	Qty = 0
	ReStock = ", DateRestocked = '"& date() &"'"
Else
	Qty = rsGetRestockItems.Fields.Item("qty").Value
	ReStock = ""
End if

set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
CommUpdate.CommandText = "UPDATE ProductDetails SET ProductDetails.qty = "& Qty &" + tbl_po_details.po_qty "+Restock+" FROM ProductDetails INNER JOIN tbl_po_details ON ProductDetails.ProductDetailID = tbl_po_details.po_detailid WHERE po_detailid ="+Request.Form("DetailID")+" AND po_orderid = "+Request("po_id")+"" 
commUpdate.Execute()

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn  
	objCmd.CommandText = "UPDATE tbl_po_details SET po_received = 1, po_date_received = '" & now() & "' WHERE po_detailid = " & Request.Form("DetailID")+ " AND po_orderid = "+Request("po_id")+"" 
	objCmd.Execute()
	Set objCmd = Nothing

ProductTitle = rsGetRestockItems.Fields.Item("title").Value
ProductGauge = rsGetRestockItems.Fields.Item("Gauge").Value
ProductLength = rsGetRestockItems.Fields.Item("Length").Value
NewQty = rsGetRestockItems.Fields.Item("po_qty").Value
Picture = rsGetRestockItems.Fields.Item("picture").Value

rsGetRestockItems.Close()
Set rsGetRestockItems = Nothing

end if ' if form is not empty %>
<html>
<head>

<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<title>Put items in stock</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../CSS/Admin.css" rel="stylesheet" type="text/css">
</head>
<body  onload="document.FRM_UpdateStock.DetailID.focus();">
<!--#include file="admin_header.asp"-->

<br>

<div style="width: 40%; padding-left: 20px">
<div class="LargeHeader">
Scan item into stock
</div>

 <div class="ContentText">
 <form ACTION="" METHOD="POST" name="FRM_UpdateStock" id="FRM_UpdateStock">
<% if Request.Form("DetailID") <> "" then %>
<span class="wishlist_text"><strong>
<img src="http://bodyartforms-products.bodyartforms.com/<%= Picture %>" alt="Image" width="50" height="50" align="absmiddle"> <% = NewQty %> put into stock for <%= ProductTitle %>&nbsp;<%= ProductGauge %> &nbsp;<%= ProductLength %></strong></span>
<% END IF %>
<p>
 Item #:
                <input type="text" name="DetailID" id="DetailID">
                <input type="submit" name="Submit" id="Submit" value="&raquo;"></td>
     
 </form>           
       

</div>
</div>
</body>
</html>

