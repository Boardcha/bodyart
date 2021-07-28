<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
	<!--#include virtual="/template/inc_includes_ajax.asp" -->
	<!--#include virtual="cart/generate_guest_id.asp"-->
	<!--#include virtual="cart/inc_cart_main.asp"-->
	<!--#include virtual="cart/inc_cart_loopitems-begin.asp"-->
	<a class="text-dark" href="/productdetails.asp?ProductID=<%= rs_getCart.Fields.Item("ProductID").Value %>">
	<div class="row px-2 py-1">
		<div class="col-auto p-0">
			<img class="img-fluid" src="https://s3.amazonaws.com/bodyartforms-products/<%= rs_getCart.Fields.Item("picture").Value %>" alt="Product photo" style="width:60px;height:auto">
		</div>
		<div class="col px-1">
			<div class="small"><%= rs_getCart.Fields.Item("mini_preview_text").Value %></div>
			<div class=" font-weight-bold text-secondary small"><%=exchange_symbol %><%= formatnumber(rs_getCart.Fields.Item("mini_line_price").Value * exchange_rate,2) %><span class="ml-3">Qty: <%= rs_getCart.Fields.Item("cart_qty").Value %></span></div>
		</div>
	</div>
</a>
	<!--#include virtual="cart/inc_cart_loopitems-end.asp"-->
	<div class="text-center font-weight-bold">	
	Subtotal: <%=exchange_symbol %><%= FormatNumber(var_subtotal * exchange_rate, 2) %>
</div>	
	<%
DataConn.Close()
Set DataConn = Nothing
%>