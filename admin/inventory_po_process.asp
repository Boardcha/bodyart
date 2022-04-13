<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% response.Buffer=false
Server.ScriptTimeout=300 %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

	var_po_id = request("po_id")

	If Request.Querystring("po_id") <> "" then
		Session("po_id") = var_po_id
	End if

' Begin updating database with submitted form information
If Request.Form("qtyadd_1") <> "" then
For i=1 to Request.Form("total")

	If Request.Form("qtyadd_" & i) <> 0 then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn  
	objCmd.CommandText = "UPDATE ProductDetails SET qty=qty + " & Request.Form("qtyadd_" & i) & ", DateRestocked = '" & date() & "' WHERE ProductDetailID = " & Request.Form("detailID_" & i)
	objCmd.Execute()
	Set objCmd = Nothing

	'Write info to edits log	
	set objCmd = Server.CreateObject("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO tbl_edits_log (user_id, detail_id, description, edit_date) VALUES (?, " & Request.Form("detailID_" & i) & ",'Automated - Added " & Request.Form("qtyadd_" & i) & " from purchase order put in stock','" & now() & "')"
	objCmd.Parameters.Append(objCmd.CreateParameter("user_id",3,1,15, rsGetUser.Fields.Item("user_id").Value ))
	objCmd.Execute()
	Set objCmd = Nothing
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn  
	objCmd.CommandText = "UPDATE tbl_po_details SET po_received = 1, po_date_received = '" & now() & "' WHERE po_detailid = " & Request.Form("po_detail_id_" & i) & " AND po_orderid = " & Session("po_id")
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



' If there are no items in the order, set it to be finalized on the current orders page

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn  
objCmd.CommandText = "SELECT po_detailid FROM tbl_po_details WHERE po_orderid = ? AND po_received = 0"
objCmd.Prepared = true
objCmd.Parameters.Append objCmd.CreateParameter("param1", 5, 1, -1, Session("po_id"))
Set TotalInOrder = objCmd.Execute		  
Set objCmd = Nothing

If TotalInOrder.EOF And TotalInOrder.BOF Then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn  
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

' build variable to break out filter to have parenthesis around the OR statement 
if Session("filter") <> "" then

var_filter_build = " AND (" & Replace(Session("filter"), "AND ", "") & ") "

end if 

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn  
objCmd.CommandText = "SELECT ProductDetails.ProductID, jewelry.title, tbl_po_details.po_orderid FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN tbl_po_details ON ProductDetails.ProductDetailID = tbl_po_details.po_detailid GROUP BY jewelry.title, tbl_po_details.po_orderid, ProductDetails.ProductID, tbl_po_details.po_received HAVING tbl_po_details.po_orderid = ? " & Session("RemoveFromMenu") & " AND tbl_po_details.po_received = 0 ORDER BY jewelry.title ASC"
objCmd.Prepared = true
objCmd.Parameters.Append objCmd.CreateParameter("param1", 5, 1, -1, Session("po_id"))
Set GetFilter = objCmd.Execute		  
Set objCmd = Nothing

' Populate drop down for remove filter
If Session("filter") <> "" then

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn  
objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.ProductID, jewelry.title, tbl_po_details.po_orderid FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN tbl_po_details ON ProductDetails.ProductDetailID = tbl_po_details.po_detailid GROUP BY jewelry.title, tbl_po_details.po_orderid, ProductDetails.ProductID HAVING (tbl_po_details.po_orderid = ?) AND (" & Session("RemoveFilter") & ") ORDER BY jewelry.title"
	objCmd.Prepared = true
	objCmd.Parameters.Append objCmd.CreateParameter("param1", 5, 1, -1, Session("po_id"))
	Set rsRemoveFilter = objCmd.Execute		  
	Set objCmd = Nothing

End if


set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn  
objCmd.CommandText = "SELECT ProductDetails.ProductDetailID, jewelry.title, ProductDetails.ProductDetail1, ProductDetails.qty, jewelry.ProductID, ProductDetails.Gauge, ProductDetails.Length, jewelry.picture, jewelry.picture_400, ProductDetails.location,  TBL_PurchaseOrders.Brand,  TBL_Barcodes_SortOrder.ID_Description, ProductDetails.BinNumber_Detail, tbl_po_details.po_qty, tbl_po_details.po_detailid, ProductDetails.price, ProductDetails.wlsl_price FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number INNER JOIN tbl_po_details ON ProductDetails.ProductDetailID = tbl_po_details.po_detailid INNER JOIN TBL_PurchaseOrders ON tbl_po_details.po_orderid = TBL_PurchaseOrders.PurchaseOrderID WHERE (tbl_po_details.po_orderid = ? AND tbl_po_details.po_qty > 0 AND tbl_po_details.po_received = 0)" & var_filter_build & " ORDER BY " & Session("SortBy")
objCmd.Prepared = true
objCmd.Parameters.Append objCmd.CreateParameter("param1", 5, 1, -1, Session("po_id"))
Set rsGetRestockItems = objCmd.Execute		  
Set objCmd = Nothing


set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn  
objCmd.CommandText = "SELECT     COUNT(ProductDetails.ProductDetailID) AS Total FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN tbl_po_details ON ProductDetails.ProductDetailID = tbl_po_details.po_detailid INNER JOIN TBL_PurchaseOrders ON tbl_po_details.po_orderid = TBL_PurchaseOrders.PurchaseOrderID WHERE (tbl_po_details.po_orderid = ? AND tbl_po_details.po_received = 0) " & var_filter_build
objCmd.Prepared = true
objCmd.Parameters.Append objCmd.CreateParameter("param1", 5, 1, -1, Session("po_id"))
Set rsTotal = objCmd.Execute		  
Set objCmd = Nothing
%>
<html>
<head>
<title>Process order</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h5>
 Process <% If Not rsGetRestockItems.EOF Or Not rsGetRestockItems.BOF Then %><%=(rsGetRestockItems.Fields.Item("brand").Value)%><% end if %> order (<%=(rsTotal.Fields.Item("total").Value)%> items displayed)
 &nbsp;&nbsp;| &nbsp;&nbsp;Purchase order #<%= Session("po_id") %>
</h5>

<div>
<% If success = "yes" Then %>
<span class="alert alert-success">Items have been updated</span>
<% end if %>
</div>

<div class="container-fluid p-0 m-0">
	<div class="row">
		<div class="col">
			<form class="form-inline" action="?po_id=<%= request("po_id") %>" method="post">
				<select class="form-control form-control-sm mr-3" name="filter" id="filter">
				<option value="0" selected>View all (default)</option>
				<% While NOT GetFilter.EOF %>  
				<option value="<%=(GetFilter.Fields.Item("ProductID").Value)%>"><%=(GetFilter.Fields.Item("title").Value)%></option>   
			<% 
			GetFilter.MoveNext()
			Wend
			%>      
				</select>
				<% If Session("filter") = "" Then %>
				<input class="btn btn-sm btn-secondary" type="submit" value="Filter">
				<% else %>
				<input class="btn btn-sm btn-secondary" type="submit" value="Add another filter">
				<% end if %>
			</form>
		</div>
		<div class="col-auto">
			<form class="form-inline" action="?po_id=<%= request("po_id") %>" method="post">
				<% If Session("filter") <> "" Then %>
				<select class="form-control form-control-sm mr-3" name="removefilter" id="removefilter">
				<% While NOT rsRemoveFilter.EOF %>
				<option value="<%=(rsRemoveFilter.Fields.Item("ProductID").Value)%>"><%=(rsRemoveFilter.Fields.Item("title").Value)%></option>
			<% 
			rsRemoveFilter.MoveNext()
			Wend
			%>
				</select>
				<input class="btn btn-sm btn-secondary" type="submit" value="Remove filter">
				<% end if %>
			</form>
		</div>
	</div><!-- row -->
</div><!-- container -->

<form class="d-block" METHOD="post" ACTION="?new=yes">
	<a class="mr-2 text-primary" href="?po_id=<%= request("po_id") %>&SortBy=ProductDetailID ASC&sort=yes">Sort by #</a>
	|
	<a class="mx-2 text-primary" href="?po_id=<%= request("po_id") %>&SortBy=title ASC&sort=yes">Sort by name</a>
	|
	<span class="ml-2 pointer text-primary" id="link-set-0">Set all qty to 0</span>

<% If Not rsGetRestockItems.EOF Or Not rsGetRestockItems.BOF Then %>

<table  class="table table-sm table-striped table-hover mt-2">
	<thead class="thead-dark">  
	<tr>
          	<th style="width:10%">Received</th>
            <th>Ordered</th>
            <th>On hand</th>
            <th>Location</th>
			<th>Retail</th>
			<th>Wholesale</th>
            <th class="Description">Description</th>
		  </tr>
		</thead>	
<% i = 0
While NOT rsGetRestockItems.EOF
i = i + 1
 %>
 <tr id="rowid-<%=(rsGetRestockItems.Fields.Item("ProductDetailID").Value)%>">
 <td class="form-inline">             
<i class="fa fa-times-circle fa-lg text-danger pointer delete-item" data-detailid="<%=(rsGetRestockItems.Fields.Item("ProductDetailID").Value)%>"></i><input class="form-control form-control-sm ml-3 check-wholesale set-0" style="width:75px" name="qtyadd_<%= i %>" type="text" id="QtyAdd" value="<%=(rsGetRestockItems.Fields.Item("po_qty").Value)%>" data-id="<%= rsGetRestockItems.Fields.Item("ProductDetailID").Value %>" />
<input name="detailID_<%= i %>" type="hidden" id="detailID" value="<%=(rsGetRestockItems.Fields.Item("ProductDetailID").Value)%>">
<input name="productID_<%= i %>" type="hidden" id="productID" value="<%=(rsGetRestockItems.Fields.Item("ProductID").Value)%>">
<input name="po_detail_id_<%= i %>" type="hidden" value="<%=(rsGetRestockItems.Fields.Item("po_detailid").Value)%>">
</td>
<td>
<%=(rsGetRestockItems.Fields.Item("po_qty").Value)%>
</td>
<td>
<%=(rsGetRestockItems.Fields.Item("qty").Value)%>
</td>
<td>
<%=(rsGetRestockItems.Fields.Item("location").Value)%> - <%=(rsGetRestockItems.Fields.Item("ID_Description").Value)%>
<% if rsGetRestockItems.Fields.Item("BinNumber_Detail").Value <> 0 then %>
			 - BIN <%=(rsGetRestockItems.Fields.Item("BinNumber_Detail").Value)%>
			<% end if %>
</td>
<td class="ajax-update">
	<input class="form-control form-control-sm check-wholesale pricecheck_retail_<%= rsGetRestockItems.Fields.Item("ProductDetailID").Value %>" name="retail_<%= rsGetRestockItems.Fields.Item("ProductDetailID").Value %>" type="text" value="<%= rsGetRestockItems.Fields.Item("price").Value %>" data-column="price" data-id="<%= rsGetRestockItems.Fields.Item("ProductDetailID").Value %>" data-friendly="Retail price">
</td>
<td class="ajax-update">
	<input class="form-control form-control-sm check-wholesale pricecheck_wlsl_<%= rsGetRestockItems.Fields.Item("ProductDetailID").Value %>" name="wlsl_<%= rsGetRestockItems.Fields.Item("ProductDetailID").Value %>" type="text" value="<%= rsGetRestockItems.Fields.Item("wlsl_price").Value %>" data-column="wlsl_price" data-id="<%= rsGetRestockItems.Fields.Item("ProductDetailID").Value %>" data-friendly="Wholesale price">
</td>
<td class="Description">
<a href="product-edit.asp?ProductID=<%=(rsGetRestockItems.Fields.Item("ProductID").Value)%>&info=less" target="_blank">
	<img src="https://bafthumbs-400.bodyartforms.com/<%=rsGetRestockItems("picture_400")%>" class="rounded float-left mr-2" style="height:50px;width:50px">
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
<input type="submit" name="button" id="button" value="FINALIZE ORDER" class="btn btn-primary">
</form>

</div><!--admin content-->
</body>
<!--#include file="includes/inc_scripts.asp"-->
<script type="text/javascript">

	//url to to do auto updating
	var auto_url = "inventory/ajax_update_inventory_view.asp"
</script>
<script type="text/javascript" src="scripts/generic_auto_update_fields.js"></script>
<script type="text/javascript">
	auto_update(); // run function to update fields when tabbing out of them

	// Delete a record
	$(document).on("click", ".delete-item", function(event){
	var po_detailid = $(this).attr("data-detailid");
	$.ajax({
		method: "POST",
		url: "inventory/ajax-delete-po-item.asp",
		data: {po_detailid:po_detailid}
		})
		.done(function(msg ) {
			$('#rowid-' + po_detailid).hide('slow');
		})
		.fail(function(msg) {
			alert('DELETE FAILED');
		});
	}); // End delete record
	
	// Throw a notice if the retail price is less than double wholesale
	$('.check-wholesale').change(function() {
        var item = $(this).attr("data-id");

		if ($('.pricecheck_retail_' + item).val() < ($('.pricecheck_wlsl_' + item).val() * 2 - .05) ) {
			$('<div class="alert alert-danger">Retail is less than double wholesale</div>').insertBefore(this).delay(10000).fadeOut();
		//	console.log($('.pricecheck_retail_' + item).val() + " , " + ($('.pricecheck_wlsl_' + item).val() * 2));
		}

		
	});

		// Set all qty to 0
		$(document).on("click", "#link-set-0", function(event){
			$('.set-0').each(function() {
					$(this).val("0");
			});
		}); // Set all qty to 0
</script>
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

DataConn.Close()
Set DataConn = Nothing
%>
