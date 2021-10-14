<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<%
check_stock = "yes"
 %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="cart/generate_guest_id.asp"-->
<!--#include virtual="cart/inc_cart_update_item.asp" -->
<% if var_orig_qty = "" then %>
{  
   "qty":"0"
}
<% elseif var_orig_qty = 0 then %>
{  
   "qty":"out of stock"
}
<% else %>
{  
   "qty":"<%= var_orig_qty %>"
}
<%
end if 


DataConn.Close()
Set DataConn = Nothing
Set rs_getCart = Nothing
%>