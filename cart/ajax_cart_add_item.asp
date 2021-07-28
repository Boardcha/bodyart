<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="cart/generate_guest_id.asp"-->
<!--#include virtual="cart/inc_cart_main.asp" -->
<!--#include virtual="cart/inc_cart_add_item.asp" -->
<%
DataConn.Close()
Set DataConn = Nothing
Set rs_getCart = Nothing
%>