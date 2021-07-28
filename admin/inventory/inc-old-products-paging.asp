<nav>
	<ul class="pagination" style="justify-content: center">
<%	if Intpage <= intpagecount then
	if intpagecount <> 1 then 
	 if Intpage <> 1 then %>
<li class="page-item"><a class="page-link text-ltpurple"  href="?action=<<&amp;intpage=<%=intpage%>"><i class="fa fa-angle-double-left fa-lg"></i></a></li>
<li class="page-item"><a class="page-link text-ltpurple"  href="?action=<&amp;intpage=<%=intpage%>"><i class="fa fa-angle-left fa-lg"></i></a></li>
<% end if %>
<li class="page-item"><span class="page-link text-white" style="background-color:#696986"><%=Intpage %></span></li>
<% if Intpage < intpagecount then %>
<li class="page-item"><a class="page-link text-ltpurple"  href="?action=>&amp;intpage=<%=intpage%>"><i class="fa fa-angle-right fa-lg"></i></a></li>
<li class="page-item"><a class="page-link text-ltpurple"  href="?action=>>&amp;intpage=<%=intpage%>"><i class="fa fa-angle-double-right fa-lg"></i> <%= intpagecount %></a></li>
<% end if 
 end if 
   end if ' if intpagecount <> 1
%>
</ul>
</nav>

