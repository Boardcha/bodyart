<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% response.Buffer=false
Server.ScriptTimeout=300 %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Set conn = Server.CreateObject ("ADODB.Connection")
conn.open = MM_bodyartforms_sql_STRING

' Begin updating database with submitted form information
If Request.Form("qtyadd_1") <> "" then
For i=1 to Request.Form("total")

	If Request.Form("qtyadd_" & i) <> 0 then
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = conn  
	objCmd.CommandText = "UPDATE ProductDetails SET qty=qty + " & Request.Form("qtyadd_" & i) & ", DateRestocked = '" & date() & "', active = 1, PurchaseOrderID = 0 WHERE ProductDetailID = " & Request.Form("detailID_" & i)
	objCmd.Execute()
	Set objCmd = Nothing
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = conn  
	objCmd.CommandText = "UPDATE jewelry SET active = 1 FROM jewelry INNER JOIN ProductDetails ON jewelry.ProductID = ProductDetails.ProductID WHERE ProductDetailID = " & Request.Form("detailID_" & i)
	objCmd.Execute()
	Set objCmd = Nothing
	
	Else
	
	' Clean up string and remove items that were qty 0 so they don't display again when the page is refreshed
	Session("filter") = Replace(Session("filter"), " OR ProductDetails.ProductID = " & Request.Form("productID_" & i), "")
	
	Session("filter") = Replace(Session("filter"), " AND ProductDetails.ProductID = " & Request.Form("productID_" & i), "")
	
	End if
	
	success = "yes"
Next
End if

' Remove item from order code
If Request.Querystring("remove") = "yes" then
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = conn  
	objCmd.CommandText = "UPDATE ProductDetails SET  PurchaseOrderID = 0 WHERE ProductDetailID = " & Request.Querystring("DetailID")
	objCmd.Execute()
	Set objCmd = Nothing
	
	remove_success = "yes"
End if

If Request.Querystring("ID") <> "" then
	Session("po_id") = Request.Querystring("ID")
End if

' If there are no items in the order, set it to be finalized on the current orders page

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = conn  
objCmd.CommandText = "SELECT ProductDetailID FROM ProductDetails WHERE PurchaseOrderID = ?"
objCmd.Prepared = true
objCmd.Parameters.Append objCmd.CreateParameter("param1", 5, 1, -1, Session("po_id"))
Set TotalInOrder = objCmd.Execute		  
Set objCmd = Nothing

If TotalInOrder.EOF And TotalInOrder.BOF Then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = conn  
	objCmd.CommandText = "UPDATE TBL_PurchaseOrders SET Received='Y', DateReceived='"& date() &"' WHERE PurchaseOrderID = " & Session("po_id")
	objCmd.Execute()
	Set objCmd = Nothing

End if

' Get sort filters
If Request.Querystring("SortBy") <> "" then
		Session("SortBy") = Request.Querystring("SortBy")
Else	
	If Session("SortBy") <> "" then
		Session("SortBy") = Session("SortBy")
	Else
		Session("SortBy") = "ProductDetailID ASC"
	End if
End if

	'Session.Abandon
	
If (Request.Querystring("new") <> "yes" AND Request.Querystring("remove") <> "yes") AND (Request.Form("filter") <> "" OR Session("filter") <> "") Then
	
	If Request.Querystring("sort") <> "yes" AND Request.Form("removefilter") = "" Then
		Session("filter") = Session("filter") & " OR ProductDetails.ProductID = " & Request.Form("filter")
	Else
	End if
Else
	If Session("filter") <> "" then
		Session("filter") = Session("filter")	
	ENd if
	If Request.Querystring("new") = "yes" then
		Session("filter") = ""
		Session("RemoveFromMenu") = ""
		Session("RemoveFilter") = ""
	End if
End if

If Request.Form("filter") = "0" then
	Session("filter") = ""
End if

'remove items from session if a filter is removed
If Request.Form("removefilter") <> "" Then

		' If when removing the LAST filter set the session status to a "view all" stype of status
	If InStr(Session("filter"), "OR") Then
	Else
		Session("filter") = ""
	ENd If
	
	Session("filter") = Replace(Session("filter"), " OR ProductDetails.ProductID = " & Request.Form("removefilter"), "")
	
	Session("filter") = Replace(Session("filter"), " AND ProductDetails.ProductID = " & Request.Form("removefilter"), "")

End if

' Get rid of first random OR in string
If InStr(Session("filter"), "AND") Then
Else
Session("filter") = Replace(Session("filter"), "OR ", "AND ", 1 , 1)
ENd If


' Replace values for remove from menu drop down based on what is selected
Session("RemoveFromMenu") = Replace(Session("filter"), "OR", "AND")
Session("RemoveFromMenu") = Replace(Session("RemoveFromMenu"), "=", "<>")

' Replace values for remove filter drop down
Session("RemoveFilter") = Replace(Session("RemoveFromMenu"), "AND", "AND-", 1 , 1)
Session("RemoveFilter") = Replace(Session("RemoveFilter"), "AND", "OR")
Session("RemoveFilter") = Replace(Session("RemoveFilter"), "OR- ", "")
Session("RemoveFilter") = Replace(Session("RemoveFilter"), "<>", "=")


'Response.write Session("filter") & "<br/>"
'Response.write Session("RemoveFromMenu") & "<br/>"
'Response.write Session("RemoveFilter")



set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = conn  
objCmd.CommandText = "SELECT ProductDetails.ProductID, jewelry.title, ProductDetails.PurchaseOrderID FROM ProductDetails INNER JOIN  jewelry ON ProductDetails.ProductID = jewelry.ProductID GROUP BY jewelry.title, ProductDetails.PurchaseOrderID, ProductDetails.ProductID HAVING (ProductDetails.PurchaseOrderID = ?) " & Session("RemoveFromMenu") & " ORDER BY title ASC"
objCmd.Prepared = true
objCmd.Parameters.Append objCmd.CreateParameter("param1", 5, 1, -1, Session("po_id"))
Set GetFilter = objCmd.Execute		  
Set objCmd = Nothing

' Populate drop down for remove filter
If Session("filter") <> "" then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = conn  
objCmd.CommandText = "SELECT ProductDetails.ProductID, jewelry.title, ProductDetails.PurchaseOrderID FROM ProductDetails INNER JOIN  jewelry ON ProductDetails.ProductID = jewelry.ProductID GROUP BY jewelry.title, ProductDetails.PurchaseOrderID, ProductDetails.ProductID HAVING (ProductDetails.PurchaseOrderID = ?) AND (" & Session("RemoveFilter") & ") ORDER BY title ASC"
	objCmd.Prepared = true
	objCmd.Parameters.Append objCmd.CreateParameter("param1", 5, 1, -1, Session("po_id"))
	Set rsRemoveFilter = objCmd.Execute		  
	Set objCmd = Nothing

End if


set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = conn  
objCmd.CommandText = "SELECT ProductDetails.ProductDetailID, jewelry.title, ProductDetails.ProductDetail1, ProductDetails.qty, ProductDetails.PurchaseOrderID, ProductDetails.POAmount, jewelry.ProductID, ProductDetails.Gauge, ProductDetails.Length, jewelry.picture, ProductDetails.location, ProductDetails.price, ProductDetails.wlsl_price, TBL_PurchaseOrders.Brand,  TBL_Barcodes_SortOrder.ID_Description FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN TBL_PurchaseOrders ON ProductDetails.PurchaseOrderID = TBL_PurchaseOrders.PurchaseOrderID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = dbo.TBL_Barcodes_SortOrder.ID_Number WHERE ProductDetails.PurchaseOrderID = ? " & Session("filter") & " ORDER BY " & Session("SortBy")
objCmd.Prepared = true
objCmd.Parameters.Append objCmd.CreateParameter("param1", 5, 1, -1, Session("po_id"))
Set rsGetRestockItems = objCmd.Execute		  
Set objCmd = Nothing


set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = conn  
objCmd.CommandText = "SELECT COUNT(ProductDetails.ProductDetailID) AS Total FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN TBL_PurchaseOrders ON ProductDetails.PurchaseOrderID = TBL_PurchaseOrders.PurchaseOrderID WHERE ProductDetails.PurchaseOrderID = ? " & Session("filter")
objCmd.Prepared = true
objCmd.Parameters.Append objCmd.CreateParameter("param1", 5, 1, -1, Session("po_id"))
Set rsTotal = objCmd.Execute		  
Set objCmd = Nothing
%>
<html>
<head>
<link href="../CSS/Admin.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/js/jquery-2.2.3.min.js"></script>
<script type="text/javascript">

	//url to to do auto updating
	var auto_url = "inventory/ajax_update-retail-wlsl-putinstock.asp"
</script>
<script type="text/javascript" src="scripts/generic_auto_update_fields.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	auto_update(); // run function
});
</script>


<title>Process order</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
.Results {
	width: 100%;
	font-size: 12px;
	color: #000;
	}
	
.Results a{
	text-decoration: none;
}

.Results tr:hover {
	background-color: #CCC;
}

.Results tr.Header td {
	background-color: #CCC;
	font-size: 14px;
	font-weight: bold;
	padding: 5px;
	}

.Results td {
	border-bottom: 1px dotted #000;
	padding: 2px;
	width: 7%;
	}

.Results td+td {
	border-left: 1px dotted #000;
	text-align: center;
	width: 7%;
	}
	
.Results td.Description {
	text-align: left;
	padding-left: 10px;
	width: 60%;
	}
	
.accepted {
		   color: #090;
		   font-weight: bold;
		   font-size: 16px;
		   }
.pricing {
	width: 4em;
}
</style>
</head>
<body>
<!--#include file="admin_header.asp"-->
<br>
 <div class="ProductsHeader">
 Process <% If Not rsGetRestockItems.EOF Or Not rsGetRestockItems.BOF Then %><%=(rsGetRestockItems.Fields.Item("brand").Value)%><% end if %> order (<%=(rsTotal.Fields.Item("total").Value)%> items displayed)
 &nbsp;&nbsp;| &nbsp;&nbsp;Purchase order #<%= Session("po_id") %></div>
<div class="ContentText">
<div>
<% If success = "yes" Then %>
<span class="accepted">Items have been updated</span>
<% end if %>
<% If remove_success = "yes" Then %>
<span class="accepted">Item has been removed from order</span>
<% end if %>
</div>
<div style="float: left;">
<form action="PurchaseOrders_PutInStock.asp" method="post">
    <select name="filter" id="filter">
      <option value="0" selected>View all (default)</option>
      <% While NOT GetFilter.EOF %>  
      <option value="<%=(GetFilter.Fields.Item("ProductID").Value)%>"><%=(GetFilter.Fields.Item("title").Value)%></option>   
  <% 
  GetFilter.MoveNext()
Wend
%>      
    </select>
    <% If Session("filter") = "" Then %>
    <input type="submit" value="Filter">
    <% else %>
    <input type="submit" value="Add another filter">
    <% end if %>
</form>
</div>
<div style="float: right;">
<form action="PurchaseOrders_PutInStock.asp" method="post">
    <% If Session("filter") <> "" Then %>
    <select name="removefilter" id="removefilter">
     <% While NOT rsRemoveFilter.EOF %>
     <option value="<%=(rsRemoveFilter.Fields.Item("ProductID").Value)%>"><%=(rsRemoveFilter.Fields.Item("title").Value)%></option>
  <% 
  rsRemoveFilter.MoveNext()
Wend
%>
    </select>
    <input type="submit" value="Remove filter">
    <% end if %>
</form>
</div>
<div style="clear: both;"></div>

<br>

<form METHOD="post" ACTION="PurchaseOrders_PutInStock.asp?new=yes">
  <input type="submit" name="button2" id="button2" value="FINALIZE ORDER">
<p><a href="PurchaseOrders_PutInStock.asp?SortBy=ProductDetailID ASC&sort=yes" class="Link_ItemDetails">Sort by #</a>&nbsp; |&nbsp;&nbsp;<a href="PurchaseOrders_PutInStock.asp?SortBy=title ASC&sort=yes" class="Link_ItemDetails">Sort by name</a></p>
<% If Not rsGetRestockItems.EOF Or Not rsGetRestockItems.BOF Then %>
<table class="Results">
          <tr class="Header">
          	<td>Received</td>
            <td>Ordered</td>
            <td>On hand</td>
			<td>Retail</td>
			<td>Wholesale</td>
            <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Location&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
            <td class="Description">Description</td>
          </tr>
<% i = 0
While NOT rsGetRestockItems.EOF
i = i + 1
 %>
 <tr>
 <td>             
<a href="PurchaseOrders_PutInStock.asp?remove=yes&DetailID=<%=(rsGetRestockItems.Fields.Item("ProductDetailID").Value)%>" style="font-weight: bold; font-size: 14px;">X</a>&nbsp;&nbsp;<input name="qtyadd_<%= i %>" type="text" id="QtyAdd" value="<%=(rsGetRestockItems.Fields.Item("POAmount").Value)%>" size="3" style="border: 1px solid #000000; padding: 1px;" />
<input name="detailID_<%= i %>" type="hidden" id="detailID" value="<%=(rsGetRestockItems.Fields.Item("ProductDetailID").Value)%>">
<input name="productID_<%= i %>" type="hidden" id="productID" value="<%=(rsGetRestockItems.Fields.Item("ProductID").Value)%>">
</td>
<td>
<%=(rsGetRestockItems.Fields.Item("POAmount").Value)%>
</td>
<td>
<%=(rsGetRestockItems.Fields.Item("qty").Value)%>
</td>
 <td class="ajax-update">             
	<input name="retail_<%= rsGetRestockItems.Fields.Item("ProductDetailID").Value %>" type="text" value="<%=(rsGetRestockItems.Fields.Item("price").Value)%>" data-column="price" data-id="<%= rsGetRestockItems.Fields.Item("ProductDetailID").Value %>" class="pricing">
</td>
 <td class="ajax-update">             
	<input name="wholesale_<%= rsGetRestockItems.Fields.Item("ProductDetailID").Value %>" type="text" value="<%=(rsGetRestockItems.Fields.Item("wlsl_price").Value)%>" data-column="wlsl_price" data-id="<%= rsGetRestockItems.Fields.Item("ProductDetailID").Value %>"  class="pricing">
</td>
<td>
<%=(rsGetRestockItems.Fields.Item("location").Value)%> - <%=(rsGetRestockItems.Fields.Item("ID_Description").Value)%></td>
<td class="Description">
<a href="product-edit.asp?ProductID=<%=(rsGetRestockItems.Fields.Item("ProductID").Value)%>&info=less" target="_blank">
<%=(rsGetRestockItems.Fields.Item("title").Value)%>&nbsp;<%=(rsGetRestockItems.Fields.Item("gauge").Value)%>&nbsp;<%=(rsGetRestockItems.Fields.Item("length").Value)%><%=(rsGetRestockItems.Fields.Item("ProductDetail1").Value)%></a>             

  </td>
  </tr>               <% 
  rsGetRestockItems.MoveNext()
Wend
%>
 </table>
  <p>
    
  </p>

          <% End If ' end Not rsGetRestockItems.EOF Or NOT rsGetRestockItems.BOF %>
     

 <% If TotalInOrder.EOF And TotalInOrder.BOF Then %>
      <span class="RequiredFields">No order to display<br>
      <br>
      Order has been completed on current orders page
      <br>
      <br>
      Go back to <a href="PurchaseOrders.asp">current orders page </a><br>
      </span>
      <% End If ' end rsGetRestockItems.EOF And rsGetRestockItems.BOF %>
<input name="total" type="hidden" id="total" value="<%= i %>">
<input type="submit" name="button" id="button" value="FINALIZE ORDER">
</form>
</div>
<p>&nbsp;</p>
</body>
</html>
<%
rsGetRestockItems.Close()
Set rsGetRestockItems = Nothing

GetFilter.Close()
Set GetFilter = Nothing

Set rsRemoveFilter = Nothing

TotalInOrder.Close()
Set TotalInOrder = Nothing

rsTotal.Close()
Set rsTotal = Nothing

conn.Close()
Set conn = Nothing
%>
