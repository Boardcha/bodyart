
<% if TotalPages > 1 then ' Only show pagination if there are more than 1 page %>
		<nav aria-label="product-paging">
				<ul class="pagination" style="justify-content: center">
									
	<% If CurrentPage > 1 then %>
	<li class="page-item"><a class="page-link text-ltpurple" href="?<%= Replace(var_qs_url, "&pagenumber=" & CurrentPage, "") %>&pagenumber=1"><i class="fa fa-chevron-double-left fa-lg"></i></a></li>
	<li class="page-item" ><a class="page-link text-ltpurple" href="?<%= Replace(var_qs_url, "&pagenumber=" & CurrentPage, "") %>&pagenumber=<%= CurrentPage - 1 %>"><i class="fa fa-chevron-left-mdc fa-lg"></i></a></li>
	<%
	End if
	%>

<%
for i = var_lowest_page to CurrentPage %>
	<% if i <> cint(CurrentPage) then %>
	<li class="page-item"><a class="page-link text-ltpurple"href="?<%= Replace(var_qs_url, "&pagenumber=" & CurrentPage, "") %>&pagenumber=<%= i %>"><%= i %></a></li>
		
	<% end if %>
<% next 
if TotalPages > 1 then
%>
<li class="page-item"><span class="page-link text-white" style="background-color:#696986"><%=CurrentPage %></span></li>

<%
end if
for i = CurrentPage to var_highest_page %>
	<% if i <> cint(CurrentPage) then %>
	<li class="page-item"><a class="page-link text-ltpurple" href="?<%= Replace(var_qs_url, "&pagenumber=" & CurrentPage, "") %>&pagenumber=<%= i %>"><%= i %></a></li>

	<% end if %>
<% next %>
	<%
	If Cint(CurrentPage) < Cint(TotalPages) then %>
	<li class="page-item"><a class="page-link text-ltpurple" href="?<%= Replace(var_qs_url, "&pagenumber=" & CurrentPage, "") %>&pagenumber=<%= CurrentPage + 1 %>"><i class="fa fa-chevron-right-mdc fa-lg"></i></a></li>
	

	<li class="page-item"><a class="page-link text-ltpurple" href="?<%= Replace(var_qs_url, "&pagenumber=" & CurrentPage, "") %>&pagenumber=<%= TotalPages %>" aria-label="Last"><i class="fa fa-chevron-double-right  fa-lg"></i> <%= TotalPages %>
		</a></li>
	<% End If %>

		</ul>
	  </nav>
	  <% end if ' Only show pagination if there are more than 1 page TotalPages > 1 %>