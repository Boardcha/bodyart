<div class="paging">
<% if Intpage <= intpagecount then %>
<% if intpagecount <> 1 then %>
<% if Intpage <> 1 then %>
<a class="alert alert-info d-inline-block p-2" href="?var_first=<%=request("var_first")%>&amp;var_last=<%=request("var_last")%>&amp;var_email=<%=request("var_email")%>&amp;action=<<&amp;intpage=<%=intpage%><%= filter_gauge_link %>&amp;custid=<%=request("custid")%>" class="admin_paging paging-style1" ><i class="fa fa-angle-double-left fa-lg"></i></a>
<a class="alert alert-info d-inline-block p-2" href="?var_first=<%=request("var_first")%>&amp;var_last=<%=request("var_last")%>&amp;var_email=<%=request("var_email")%>&amp;action=<&amp;intpage=<%=intpage%><%= filter_gauge_link %>&amp;custid=<%=request("custid")%>" class="admin_paging paging-style1" ><i class="fa fa-angle-left fa-lg"></i></a>
<% end if %>
Page: <%=Intpage & " of " & intpagecount%>
<% if Intpage < intpagecount then %>
<a class="alert alert-info d-inline-block p-2" href="?var_first=<%=request("var_first")%>&amp;var_last=<%=request("var_last")%>&amp;var_email=<%=request("var_email")%>&amp;action=>&amp;intpage=<%=intpage%><%= filter_gauge_link %>&amp;custid=<%=request("custid")%>" class="admin_paging paging-style1" ><i class="fa fa-angle-right fa-lg"></i></a>
<a class="alert alert-info d-inline-block p-2" href="?var_first=<%=request("var_first")%>&amp;var_last=<%=request("var_last")%>&amp;var_email=<%=request("var_email")%>&amp;action=>>&amp;intpage=<%=intpage%><%= filter_gauge_link %>&amp;custid=<%=request("custid")%>" class="admin_paging paging-style1" ><i class="fa fa-angle-double-right fa-lg"></i></a>
<% end if %>
<% end if %>
<%
   end if ' if intpagecount <> 1
%>
</div>