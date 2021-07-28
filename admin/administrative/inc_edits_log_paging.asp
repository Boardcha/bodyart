<div class="admin_paging_div paging paging-div my-3 text-center">
<% if Intpage <= intpagecount then %>
<% if intpagecount <> 1 then %>
<input type="hidden" name="intpage" value="<%=intpage%>">
<% if Intpage <> 1 then %>
<a href="?action=<<&amp;intpage=<%=intpage%>&amp;user=<%= request("user") %>&amp;productid=<%= request("productid") %>&amp;detailid=<%= request("detailid") %>&amp;refunds=<%= request("refunds") %>" class="btn btn-sm btn-info" ><i class="fa fa-angle-double-left fa-lg"></i></a>
<a href="?action=<&amp;intpage=<%=intpage%>&amp;user=<%= request("user") %>&amp;productid=<%= request("productid") %>&amp;detailid=<%= request("detailid") %>&amp;refunds=<%= request("refunds") %>" class="btn btn-sm btn-info" ><i class="fa fa-angle-left fa-lg"></i></a>
<% end if %>
Page: <%=Intpage & " of " & intpagecount%>
<% if Intpage < intpagecount then %>
<a href="?action=>&amp;intpage=<%=intpage%>&amp;user=<%= request("user") %>&amp;productid=<%= request("productid") %>&amp;detailid=<%= request("detailid") %>&amp;refunds=<%= request("refunds") %>" class="btn btn-sm btn-info" ><i class="fa fa-angle-right fa-lg"></i></a>
<a href="?action=>>&amp;intpage=<%=intpage%>&amp;user=<%= request("user") %>&amp;productid=<%= request("productid") %>&amp;detailid=<%= request("detailid") %>&amp;refunds=<%= request("refunds") %>" class="btn btn-sm btn-info" ><i class="fa fa-angle-double-right fa-lg"></i></a>
<% end if %>
<% end if 
   end if ' if intpagecount <> 1
%>
</div>