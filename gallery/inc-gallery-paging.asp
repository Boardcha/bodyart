<nav class="text-center">
   <%if intpagecount > 1 then 'only show if there's more than 1 page to go through %>
     <div class="small text-secondary mb-1">Displaying 20 photos per page</div>  
     <% end if %> 
    <ul class="pagination" style="justify-content: center">
<% if Intpage <= intpagecount then %>
<% if intpagecount <> 1 then %>
<% if Intpage <> 1 then %>
<li class="page-item"><a class="page-link text-ltpurple" href="?ProductID=<%=request("ProductID")%>&amp;action=<<&amp;intpage=<%=intpage%><%= filter_gauge_link %><%= filter_color_link %>"><i class="fa fa-chevron-double-left fa-lg"></i></a></li>
<li class="page-item"><a class="page-link text-ltpurple" href="?ProductID=<%=request("ProductID")%>&amp;action=<&amp;intpage=<%=intpage%><%= filter_gauge_link %><%= filter_color_link %>"><i class="fa fa-chevron-left-mdc fa-lg"></i></a></li>
<% end if %>
<li class="page-item"><span class="page-link text-white" style="background-color:#696986"><%=Intpage %></span></li>
<% if Intpage < intpagecount then %>
<li class="page-item"><a class="page-link text-ltpurple" href="?ProductID=<%=request("ProductID")%>&amp;action=>&amp;intpage=<%=intpage%><%= filter_gauge_link %><%= filter_color_link %>"><i class="fa fa-chevron-right-mdc fa-lg"></i></a></li>
<li class="page-item"><a class="page-link text-ltpurple" href="?ProductID=<%=request("ProductID")%>&amp;action=>>&amp;intpage=<%=intpage%><%= filter_gauge_link %><%= filter_color_link %>"><i class="fa fa-chevron-double-right fa-lg"></i> <%= intpagecount%></a></li>
<% end if %>
<% end if 
   end if ' if intpagecount <> 1
%>
</ul>
</nav>