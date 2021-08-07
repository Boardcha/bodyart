<%@ Language=VBScript %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<html>
<head>
<!--#include file="includes/inc_scripts.asp"-->
<title>
	Review products
</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">

<table class="table table-hover">
	<thead class="thead-dark">
		<tr>
			<th>Title</th>
            <th>Status</th>
			<th>Date Created</th>
			<th></th>
		</tr>
	</thead>	
<%
		While NOT rsGetProductsToReview.EOF
%>
		<tr>
			<td>
                <a href="product-edit.asp?ProductID=<%= rsGetProductsToReview("ProductID") %>" target="_blank">
				<img class="mr-3" src="http://bodyartforms-products.bodyartforms.com/<%= rsGetProductsToReview("picture") %>" width="90" height="90">
                <%= rsGetProductsToReview("title") %>
                </a>
			</td>
            <td class="align-middle">
                <% if rsGetProductsToReview("active") = 1 then %>
                    <span class="alert alert-success font-weight-bold">Active</span>
                <% else %>
                    <span class="alert alert-danger font-weight-bold">Inactive</span>
                <% end if %>
            </td>
			<td class="align-middle">
				<%= rsGetProductsToReview("date_added") %>
			</td>
			<td class="align-middle">
				Added by <%= rsGetProductsToReview("added_by") %>
				<% if rsGetProductsToReview("reviewed_by_1") <> "" then %>
				<div>
				Reviewed by <%= rsGetProductsToReview("reviewed_by_1") %><span class="ml-2"><%= rsGetProductsToReview("review_date_1") %></span>
				</div>
			<% end if %>
			<% if rsGetProductsToReview("reviewed_by_2") <> "" then %>
				<div>
				Reviewed by <%= rsGetProductsToReview("reviewed_by_2") %><span class="ml-2"><%= rsGetProductsToReview("review_date_2") %></span>
				</div>
			<% end if %>
			</td>
		</tr>
<% 
	rsGetProductsToReview.MoveNext()
	Wend
%>
</table>

</div>
</body>
</html>
<%
DataConn.Close()
Set DataConn = Nothing
%>
