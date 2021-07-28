<%@LANGUAGE="VBSCRIPT" %>
<% response.Buffer = false %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


if request("user") <> "" then	
	sql_user = " AND e.user_id = ?"
else
	sql_user = ""
end if

if request("productid") <> "" then
	sql_product = " AND e.product_id = ?"
else	
	sql_product = ""
end if

if request("detailid") <> "" then
	sql_detail = " AND e.detail_id = ?"
else
	sql_detail = ""
end if

if request("invoiceid") <> "" then
	sql_invoice = " AND e.invoice_id = ?"
else
	sql_invoice = ""
end if

if request.querystring("refunds") = "yes" then
	sql_refunds = " AND (e.description like '%account credit%' OR e.description like '%Refunded%')"
end if

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT top(1000) e.description, e.edit_date, u.name, j.title, e.invoice_id, d.Gauge, d.Length, d.ProductDetail1, j.picture, j.ProductID, e.detail_id, e.invoice_detail_id, e.product_id, e.user_id FROM TBL_AdminUsers AS u INNER JOIN tbl_edits_log AS e ON u.ID = e.user_id LEFT OUTER JOIN TBL_OrderSummary AS s ON e.invoice_detail_id = s.OrderDetailID LEFT OUTER JOIN sent_items AS i ON e.invoice_id = i.ID LEFT OUTER JOIN jewelry AS j ON e.product_id = j.ProductID LEFT OUTER JOIN ProductDetails AS d ON e.detail_id = d.ProductDetailID WHERE (e.edit_date <> '') " & sql_product & " " & sql_detail & " " & sql_invoice & " " & sql_user & " " & sql_refunds & " ORDER BY e.edit_date DESC"
	if request("productid") <> "" then
		objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,10,request("productid")))
	end if
	if request("detailid") <> "" then
		objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10,request("detailid")))
	end if
	if request("invoiceid") <> "" then
		objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10,request("invoiceid")))
	end if
	if request("user") <> "" then
		objCmd.Parameters.Append(objCmd.CreateParameter("userid",3,1,10,request("user")))
	end if

	
	set rs_getEdits = Server.CreateObject("ADODB.Recordset")
	rs_getEdits.CursorLocation = 3 'adUseClient
	rs_getEdits.Open objCmd
	rs_getEdits.PageSize = 50 ' not using (possibly needed for pagination)
	intPageCount = rs_getEdits.PageCount ' not using (possibly needed for pagination)


	Select Case Request("Action")
		case "<<"
			intpage = 1
		case "<"
			intpage = Request("intpage")-1
			if intpage < 1 then intpage = 1
		case ">"
			intpage = Request("intpage")+1
			if intpage > intPageCount then intpage = IntPageCount
		Case ">>"
			intpage = intPageCount
		case else
			intpage = 1
	end select
	
	
	' Get users for select menu
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT ID, name FROM TBL_AdminUsers WHERE archived = 0  ORDER BY name ASC"
	Set rs_getAdminUsers = objCmd.Execute()
%>
<!DOCTYPE html> 
<html>
<head>
<title>Edit log search</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<form class="form-inline mb-2" action="edits_logs.asp" method="POST">
	    
	<input class="form-control form-control-sm mr-2" type="text" name="productid" placeholder="Product ID #" value="<%= request("productid") %>">
	
	<input class="form-control form-control-sm mr-2" type="text" name="detailid" placeholder="Detail ID #" value="<%= request("detailid") %>">
	
	<input class="form-control form-control-sm mr-2" type="text" name="invoiceid" placeholder="Invoice #" value="<%= request("invoiceid") %>">
	
	<select class="form-control form-control-sm mr-2" name="user">
		<option value="">Select employee</option>
<%
		While NOT rs_getAdminUsers.EOF
%>
		<option value="<%= rs_getAdminUsers.Fields.Item("ID").Value %>"><%= rs_getAdminUsers.Fields.Item("name").Value %></option>
<% 
	rs_getAdminUsers.MoveNext()
	Wend
%>
	</select>
	
	<button class="btn btn-sm btn-purple" type="submit">Submit</button>	
</form>
<a href="?refunds=yes">Store credits / Refunds</a>

<% if NOT rs_getEdits.EOF then ' only show if recordset has results %>
<!--#include file="administrative/inc_edits_log_paging.asp" -->
<table class="table table-striped table-borderless table-hover">
	<thead class="thead-dark">
		<tr>
			<th width="10%">Edited by</th>
			<th width="70%">Description</th>
			<th width="20%">Date edited</th>
		</tr>
	</thead>
<%
'	While NOT rs_getEdits.EOF
rs_getEdits.AbsolutePage = intPage '======== PAGING
For intRecord = 1 To rs_getEdits.PageSize 
%>
<tbody>
	<tr>

		<td>
			<%= rs_getEdits.Fields.Item("name").Value %>
		</td>
		<td>
		<% 
			if rs_getEdits.Fields.Item("ProductID").Value <> 0 then %>
			<a href="product-edit.asp?ProductID=<%= rs_getEdits.Fields.Item("ProductID").Value %>" target="_blank"><img src="http://bodyartforms-products.bodyartforms.com/<%=(rs_getEdits.Fields.Item("picture").Value)%>" width="40px" style="padding-right: 10px;" align="left"></a><%= rs_getEdits.Fields.Item("title").Value %>
		<% end if %>
	
			<% 
			if rs_getEdits.Fields.Item("detail_id").Value <> 0 then %>
			
			&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp; Item: <%= rs_getEdits.Fields.Item("Gauge").Value %>&nbsp;<%= rs_getEdits.Fields.Item("Length").Value %>&nbsp;<%= rs_getEdits.Fields.Item("ProductDetail1").Value %>
			
			<% end if ' if detail id <> 0 %>
			<% 
			if rs_getEdits.Fields.Item("invoice_id").Value <> 0 then %>
			
			Invoice # <a href="invoice.asp?ID=<%= rs_getEdits.Fields.Item("invoice_id").Value %>" target="_blank"><%= rs_getEdits.Fields.Item("invoice_id").Value %></a>
			<% end if ' if detail id <> 0 %>
			<br/><br/>
			<strong>
			<%= rs_getEdits.Fields.Item("description").Value %></strong>
		</td>
		<td>
			<%= rs_getEdits.Fields.Item("edit_date").Value %>	
		</td>
	</tr>
</tbody>

<%
	rs_getEdits.MoveNext()
	
	If rs_getEdits.EOF Then Exit For  ' ====== PAGING
Next ' ====== PAGING
'Wend
%>
</table>
<!--#include file="administrative/inc_edits_log_paging.asp" -->
<% else ' if no records found %>
No edits found
<% end if ' if recordset has results to show %>
</div><!-- end content area div -->
</body>
</html>
<%
Set rs_getEdits = Nothing
set rs_getAdminUsers = nothing
DataConn.Close()
%>
