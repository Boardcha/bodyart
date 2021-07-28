<%
' START Display stock levels notice if needed
%>
<div id="stock_notice" class="stock_notification">
</div>
<% if stock_display <> "" or session("stock_display") <> "" then
		stock_notice_show = ""
	else
		stock_notice_show = "display:none"
	end if
%>
<span style="<%= stock_notice_show %>">
	<div class="stock_notification">
		<div class="alert alert-danger alert-dismissible fade show mt-2" role="alert">
				Unfortunately we do not have enough quantity in stock of the item(s) you wanted below. We have adjusted your quantity to what we have available.
				<blockquote>
						<%= session("stock_display") %>
				</blockquote>
				<button type="button" class="close" data-dismiss="alert" aria-label="Close">
					<span aria-hidden="true">&times;</span>
				</button>
				</div>
    </div><!-- END DIV row -->	
</span>
<%' END stock levels notice
%>