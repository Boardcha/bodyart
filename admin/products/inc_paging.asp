<% if Intpage <= intpagecount then %>
<% if intpagecount <> 1 then %>
<input type="hidden" class="btn btn-sm btn-info" name="intpage" value="<%=intpage%>">
<% if Intpage <> 1 then %>
<a href="?ProductID=<%=request("ProductID")%>&amp;action=<<&amp;intpage=<%=intpage%>&amp;filter_gauge=<%= request("filter_gauge") %>&amp;filter_active=<%= request("filter_active") %>&amp;filter_detailid=<%= request("filter_id") %>" class="btn btn-sm btn-info" ><i class="fa fa-angle-double-left fa-lg text-white"></i></a>
<a href="?ProductID=<%=request("ProductID")%>&amp;action=<&amp;intpage=<%=intpage%>&amp;filter_gauge=<%= request("filter_gauge") %>&amp;filter_active=<%= request("filter_active") %>&amp;filter_detailid=<%= request("filter_id") %>" class="btn btn-sm btn-info" ><i class="fa fa-angle-left fa-lg text-white"></i></a>
<% end if %>
<span class="mx-2">Page <%=Intpage & " of " & intpagecount%></span>
<% if Intpage < intpagecount then %>
<a href="?ProductID=<%=request("ProductID")%>&amp;action=>&amp;intpage=<%=intpage%>&amp;filter_gauge=<%= request("filter_gauge") %>&amp;filter_active=<%= request("filter_active") %>&amp;filter_detailid=<%= request("filter_id") %>" class="btn btn-sm btn-info" ><i class="fa fa-angle-right fa-lg text-white"></i></a>
<a href="?ProductID=<%=request("ProductID")%>&amp;action=>>&amp;intpage=<%=intpage%>&amp;filter_gauge=<%= request("filter_gauge") %>&amp;filter_active=<%= request("filter_active") %>&amp;filter_detailid=<%= request("filter_id") %>" class="btn btn-sm btn-info" ><i class="fa fa-angle-double-right fa-lg text-white"></i></a>
<% end if %>
<% end if 
   end if ' if intpagecount <> 1
%>