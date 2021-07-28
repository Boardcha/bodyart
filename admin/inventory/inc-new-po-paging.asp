<% if Intpage <= intpagecount then %>
<% if intpagecount <> 1 then %>
<% if Intpage <> 1 then %>
<a href="?brand=<%=request("brand")%>&amp;resume=yes&amp;action=<<&amp;intpage=<%=intpage%>&amp;autoclave=<%= request.querystring("autoclave") %>&amp;keywords_title=<%= request.querystring("keywords_title") %>&amp;keywords_details=<%= request.querystring("keywords_details") %>" class="btn btn-sm btn-info" ><i class="fa fa-angle-double-left fa-lg"></i></a>
<a href="?brand=<%=request("brand")%>&amp;resume=yes&amp;action=<&amp;intpage=<%=intpage%>&amp;autoclave=<%= request.querystring("autoclave") %>&amp;keywords_title=<%= request.querystring("keywords_title") %>&amp;keywords_details=<%= request.querystring("keywords_details") %>" class="paging-reviews btn btn-sm btn-info" ><i class="fa fa-angle-left fa-lg"></i></a>
<% end if %>
<%=Intpage & " of " & intpagecount%>
<% if Intpage < intpagecount then %>
<a href="?brand=<%=request("brand")%>&amp;resume=yes&amp;action=>&amp;intpage=<%=intpage%>&amp;autoclave=<%= request.querystring("autoclave") %>&amp;keywords_title=<%= request.querystring("keywords_title") %>&amp;keywords_details=<%= request.querystring("keywords_details") %>" class="paging-reviews btn btn-sm btn-info" ><i class="fa fa-angle-right fa-lg"></i></a>
<a href="?brand=<%=request("brand")%>&amp;resume=yes&amp;action=>>&amp;intpage=<%=intpage%>&amp;autoclave=<%= request.querystring("autoclave") %>&amp;keywords_title=<%= request.querystring("keywords_title") %>&amp;keywords_details=<%= request.querystring("keywords_details") %>" class="paging-reviews btn btn-sm btn-info" ><i class="fa fa-angle-double-right fa-lg"></i></a>
<% end if %>
<% end if 
   end if ' if intpagecount <> 1
%>