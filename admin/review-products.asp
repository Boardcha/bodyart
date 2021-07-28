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
		While NOT rsGetProducts.EOF
%>
		<tr>
			<td>
                <a href="product-edit.asp?ProductID=<%= rsGetProducts("ProductID") %>" target="_blank">
				<img class="mr-3" src="http://bodyartforms-products.bodyartforms.com/<%= rsGetProducts("picture") %>" width="90" height="90">
                <%= rsGetProducts("title") %>
                </a>
			</td>
            <td class="align-middle">
                <% if rsGetProducts("active") = 1 then %>
                    <span class="alert alert-success font-weight-bold">Active</span>
                <% else %>
                    <span class="alert alert-danger font-weight-bold">Inactive</span>
                <% end if %>
            </td>
			<td class="align-middle">
				<%= rsGetProducts("date_added") %>
			</td>
			<td class="align-middle">
				Added by <%= rsGetProducts("added_by") %>
				<% if rsGetProducts("reviewed_by_1") <> "" then %>
				<div>
				Reviewed by <%= rsGetProducts("reviewed_by_1") %><span class="ml-2"><%= rsGetProducts("review_date_1") %></span>
				</div>
			<% end if %>
			<% if rsGetProducts("reviewed_by_2") <> "" then %>
				<div>
				Reviewed by <%= rsGetProducts("reviewed_by_2") %><span class="ml-2"><%= rsGetProducts("review_date_2") %></span>
				</div>
			<% end if %>
			</td>
		</tr>
<% 
	rsGetProducts.MoveNext()
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
