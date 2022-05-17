<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="cart/generate_guest_id.asp"-->
<%
var_sql_query = " AND customorder = 'yes'"
%>
<!--#include virtual="cart/inc_cart_main.asp" -->
<div class="alert alert-warning p-2">
	<strong>Your order contains custom made items</strong>
	<br>
	Your ENTIRE ORDER will be held until the all the custom pieces arrive to ship to you. If you would like to receive the items that are not custom first, please place a separate order for those items.
</div>
<%
rs_getCart.ReQuery() 
While Not rs_getCart.eof
%>
<div style="float:left; clear: both">
	<img class="float-left mr-2 mb-1"  src="https://s3.amazonaws.com/bodyartforms-products/<%= rs_getCart("picture") %>" alt="Product photo">
	<%= rs_getCart("title") %>&nbsp;<%=(rs_getCart("gauge"))%>&nbsp;<%=(rs_getCart("length"))%>&nbsp;<%=(rs_getCart("ProductDetail1"))%>&nbsp;<%=Sanitize(rs_getCart("cart_preorderNotes"))%>
	<span class="d-inline-block my-1 alert alert-warning px-2 py-1">
		<%= rs_getCart("preorder_timeframes") %> to receive
	</span>	
</div>
<%
rs_getCart.MoveNext
Wend

DataConn.Close()
Set DataConn = Nothing
%>