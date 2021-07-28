<%@LANGUAGE="VBSCRIPT" %>
<%Response.Buffer = False%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	
	objCmd.CommandText = "SELECT x.email, x.customer_ID, x.credits, CONVERT(varchar, x.last_login, 101) AS last_login, CONVERT(varchar, x.account_created, 101) AS account_created FROM customers AS x INNER JOIN (SELECT email FROM customers AS t GROUP BY email HAVING (COUNT(email) > 1)) AS y ON y.email = x.email ORDER BY x.email ASC, x.last_login DESC, x.customer_ID DESC"

	set rs_getcustomer = Server.CreateObject("ADODB.Recordset")
	rs_getcustomer.CursorLocation = 3 'adUseClient
	rs_getcustomer.Open objCmd
%>
<html>
<head>
<title>Duplicate Customer Accounts</title>
<style>
.span-button, .match {font-size: .9em;}
.span-button:hover, .match:hover {
	filter: brightness(85%);
}
</style>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="content-grey">

	
	
<% if NOT rs_getcustomer.EOF then

	TotalRecords = rs_getcustomer.RecordCount
	rs_getcustomer.PageSize = 500
	TotalPages = rs_getcustomer.PageCount
	
	if Request.Querystring("pagenumber") = "" then
		CurrentPage = 1
	else
		temp_pagenumber = cint(Request.Querystring("pagenumber"))
	end if

	If temp_pagenumber = 0 or temp_pagenumber > cint(TotalPages) Then
		CurrentPage = 1
	Else
		CurrentPage = Request.Querystring("pagenumber")
	End if
end if
%>

<br/>
<% if NOT rs_getcustomer.EOF then %>
<%
' if more than 4 pages show ... for last page
if TotalPages > 4 then
	var_total_pages = "... " & "<a href=""?pagenumber=" & TotalPages & """ class=""ind-page"">" & TotalPages & "</a>"
end if

' Retrieve page numbers BEFORE current page
	' Get current page - 3
	var_lowest_page = CurrentPage - 3
	' If lowest page is below 0, set the page number to 1
	if var_lowest_page <= 0 then
		var_lowest_page = 1
	end if

' Retrieve page numbers AFTER current page
	' Get current page + 3
	var_highest_page = CurrentPage + 3
	' If highest page is above the total pages, then set it to the total pages
	if var_highest_page >= TotalPages then
		var_highest_page = TotalPages
		var_total_pages = "" ' don't show elipses in paging
	end if

	var_begin_results_count = (CurrentPage - 1) * session("resultsperpage") + 1
	var_end_results_count = CurrentPage * session("resultsperpage")
	%>


<span class="product-paging-links">	
	<% If CurrentPage > 1 then %>
	<a href="assigning-detail-attributes.asp" class="paging-style1"><i class="fa fa-angle-double-left fa-lg"></i></a> 
	<a href="?pagenumber=<%= CurrentPage - 1 %>" class="paging-style1"> <i class="fa fa-angle-left fa-lg"></i></a>
	<%
	End if
	%>
<%
for i = var_lowest_page to CurrentPage %>
	<% if i <> cint(CurrentPage) then %>
		<a href="?pagenumber=<%= i %>" class="ind-page"><%= i %></a>
	<% end if %>
<% next 
if TotalPages > 1 then
%>
<span class="current-page"><%= CurrentPage %></span>
<%
end if
for i = CurrentPage to var_highest_page %>
	<% if i <> cint(CurrentPage) then %>
		<a href="?pagenumber=<%= i %>" class="ind-page"><%= i %></a>
	<% end if %>
<% next %>
<%= var_total_pages %>	
	<%
	'response.write "<b>Current Page: " & CurrentPage & " AND Total Pages: " & TotalPages & "</b>"
	If Cint(CurrentPage) < Cint(TotalPages) then %>
	<a href="?pagenumber=<%= CurrentPage + 1 %>" class="paging-style1"><i class="fa fa-angle-right fa-lg"></i></a> 
	
	<a href="?pagenumber=<%= TotalPages %>" class="paging-style1"><i class="fa fa-angle-double-right fa-lg"></i></a>
	<% End If %>
</span>	
</div>

<table class="admin-table attr-fields ajax-update" style="width: 100%; border-spacing: 0">
	<thead>
		<th>Email</th>
		<th>ID</th>
		<th>Last login</th>
		<th>Account created</th>
		<th>Store credit</th>
	</thead>

<style>
	.admin-table select{font-size: 1em; padding: .2em; margin: 0 .5em 0 0}
	.divgroup{padding-top:1em}
</style>
<%


	item_number = 0
	rs_getcustomer.AbsolutePage = CurrentPage '======== PAGING
	For intRecord = 1 To rs_getcustomer.PageSize
	item_number = item_number + 1
%>
	<tbody class="admin-tbody-border-bottom row-group">
		<tr class="no-border">
			<td style="width:5%">
				<%= rs_getcustomer.Fields.Item("email").Value %>
			</td>
			<td style="width:5%">
				<%= rs_getcustomer.Fields.Item("customer_ID").Value %>
			</td>
			<td style="width:5%">
				<%= rs_getcustomer.Fields.Item("last_login").Value %>
			</td>
			<td style="width:5%">
				<%= rs_getcustomer.Fields.Item("account_created").Value %>
			</td>
			<td style="width:80%">
				<% if rs_getcustomer.Fields.Item("credits").Value <> 0 then %>
				<%= formatcurrency(rs_getcustomer.Fields.Item("credits").Value,2) %>
				<% end if %>
			</td>			
		</tr>
		<tr>
			<td colspan=5>
				<span class="span-button status" style="background-color:#F78181;padding:.4em;border-radius:.4em;margin:0 .5em;cursor:pointer" data-status="delete" data-id="<%= rs_getcustomer.Fields.Item("customer_ID").Value %>">Delete</span>
				&nbsp;&nbsp;&nbsp;&nbsp;
			<span class="span-button status" style="background-color:#5FB404;padding:.4em;border-radius:.4em;margin:0 .5em;cursor:pointer" data-status="keep" data-id="<%= rs_getcustomer.Fields.Item("customer_ID").Value %>">Keep</span>
			</td>
		</tr>
	</tbody>
<% 
  rs_getcustomer.MoveNext()
If rs_getcustomer.EOF Then Exit For  ' ====== PAGING
Next ' ====== PAGING

end if ' if recordset not empty
%>
	</table>

	<br/><br/>
	
<div class="paging" style="clear:both">

<%
' if more than 4 pages show ... for last page
if TotalPages > 4 then
	var_total_pages = "... " & "<a href=""?pagenumber=" & TotalPages & """ class=""ind-page"">" & TotalPages & "</a>"
end if

' Retrieve page numbers BEFORE current page
	' Get current page - 3
	var_lowest_page = CurrentPage - 3
	' If lowest page is below 0, set the page number to 1
	if var_lowest_page <= 0 then
		var_lowest_page = 1
	end if

' Retrieve page numbers AFTER current page
	' Get current page + 3
	var_highest_page = CurrentPage + 3
	' If highest page is above the total pages, then set it to the total pages
	if var_highest_page >= TotalPages then
		var_highest_page = TotalPages
		var_total_pages = "" ' don't show elipses in paging
	end if

	var_begin_results_count = (CurrentPage - 1) * session("resultsperpage") + 1
	var_end_results_count = CurrentPage * session("resultsperpage")
	%>


<span class="product-paging-links">	
	<% If CurrentPage > 1 then %>
	<a href="assigning-detail-attributes.asp" class="paging-style1"><i class="fa fa-angle-double-left fa-lg"></i></a> 
	<a href="?pagenumber=<%= CurrentPage - 1 %>" class="paging-style1"> <i class="fa fa-angle-left fa-lg"></i></a>
	<%
	End if
	%>
<%
for i = var_lowest_page to CurrentPage %>
	<% if i <> cint(CurrentPage) then %>
		<a href="?pagenumber=<%= i %>" class="ind-page"><%= i %></a>
	<% end if %>
<% next 
if TotalPages > 1 then
%>
<span class="current-page"><%= CurrentPage %></span>
<%
end if
for i = CurrentPage to var_highest_page %>
	<% if i <> cint(CurrentPage) then %>
		<a href="?pagenumber=<%= i %>" class="ind-page"><%= i %></a>
	<% end if %>
<% next %>
<%= var_total_pages %>	
	<%
	'response.write "<b>Current Page: " & CurrentPage & " AND Total Pages: " & TotalPages & "</b>"
	If Cint(CurrentPage) < Cint(TotalPages) then %>
	<a href="?pagenumber=<%= CurrentPage + 1 %>" class="paging-style1"><i class="fa fa-angle-right fa-lg"></i></a> 
	
	<a href="?pagenumber=<%= TotalPages %>" class="paging-style1"><i class="fa fa-angle-double-right fa-lg"></i></a>
	<% End If %>
</span>	
</div>
	</div>
<!-- end main product edit section / grey background div -->
</body>
</html>

<script type="text/javascript">
	$(document).ready(function(){
	
		
		$(".status").click(function(){
			var status = $(this).attr("data-status");
			var id = $(this).attr("data-id");
					
			$.ajax({
				method: "POST",
				url: "temp_projects/ajax-duplicate-customers.asp",
				data: {status: status, id: id}
			})
				.done(function( msg ) {
					console.log("Success");
			})
				.fail(function(msg) {
					alert("Update failed");
			});
			
				$(this).closest('tbody').find('td').fadeOut('fast', 
			function(here){ 
				$(here).parents('tr:first').remove();                    
			});    
	
		return false;
		}); // end click function
	
		
	}); // end document ready
	</script>
<%
DataConn.Close()
%>