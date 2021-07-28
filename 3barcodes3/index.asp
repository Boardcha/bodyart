<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set objcmd = Server.CreateObject("ADODB.command")
objcmd.ActiveConnection = DataConn
objcmd.CommandText = "SELECT Count(*) AS total_products FROM jewelry WHERE to_be_pulled = 1 AND pull_completed = 0"
Set rsGetDiscontinued = objcmd.Execute()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0" />
<meta name="mobile-web-app-capable" content="yes">
<title>Barcode management</title>
<link href="/CSS/baf.min.css?v=092519" rel="stylesheet" type="text/css" />
</head>

<body>
	 <!--#include file="includes/scanners-header.asp" -->
<div class="p-3">
	 <h6>Invoices</h6>
<ul class="list-group list-group-flush">
		<li class="list-group-item"><a href="pulling-invoices.asp">Scan invoice</a></li>
		<li class="list-group-item"><a href="review-backorders.asp">Review backorders</a></li>
		</ul>
<% if rsGetDiscontinued.Fields.Item("total_products").Value > 0 then%>
<h6 class="mt-4">Pull Products</h6>
<ul class="list-group list-group-flush">
		<li class="list-group-item"><a href="pull-discontinued-products.asp">Pull discontinued products</a><span class="badge badge-danger ml-2"><%= rsGetDiscontinued.Fields.Item("total_products").Value %></span></li>
		</ul>
		<% end if %>
		<h6 class="mt-4">Restock & Tag Products</h6>
<ul class="list-group list-group-flush">
	<li class="list-group-item"><a href="tag-items.asp">Tag products</a></li>
<li class="list-group-item"><a href="restock-items.asp">Restock products</a></li>
</ul>
<h6 class="mt-4">Assigning Products</h6>
<ul class="list-group list-group-flush">
<li class="list-group-item"><a href="barcode_convert.asp">Scan items into sections</a></li>
<li class="list-group-item"><a href="assign-discontinued-items.asp">Assign discontinued items</a></li>
<li class="list-group-item"><a href="scan_front-cases.asp">Scan into front cases</a></li>
<li class="list-group-item"><a href="barcode_assignloc.asp">Assign regular items to location &amp; section</a></li>
</ul>
<h6 class="mt-4">Inventory Counts</h6>
<ul class="list-group list-group-flush">
<li class="list-group-item"><a href="ItemCount.asp">Inventory count (regular item)</a></li>
<li class="list-group-item"><a href="/admin/inventory-count-limited-bin.asp">Inventory count (limited bin)</a></li>
</ul>

</div>
</body>
</html>
