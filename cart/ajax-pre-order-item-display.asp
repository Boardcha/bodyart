<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="cart/generate_guest_id.asp"-->
<div class="alert alert-warning p-2">
	<strong>Your order contains custom made (PRE-ORDER) items.</strong>
	<br>
	Your ENTIRE ORDER will be held until the custom piece arrives to ship to you.
</div>
<%
If request.cookies("OrderAddonsActive") <> "" Then
	var_addons = " AND cart_addon_item = 1 "
Else
	var_addons = " AND cart_addon_item = 0 "
End if

' -- RETRIEVE SHOPPING CART CONTENTS --
Set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT tbl_carts.cart_id, tbl_carts.cart_detailID, tbl_carts.cart_preorderNotes, tbl_carts.cart_custId, tbl_carts.cart_qty, tbl_carts.cart_wishlistid, jewelry.title, jewelry.internal, jewelry.customorder, jewelry.autoclavable, jewelry.brandname, jewelry.SaleDiscount, jewelry.secret_sale, jewelry.ProductID, jewelry.pair, jewelry.picture, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.price, ProductDetails.wlsl_price, ProductDetails.qty, ProductDetails.ProductDetail1, ProductDetails.ProductDetailID, ProductDetails.free, jewelry.SaleExempt, jewelry.jewelry, TBL_Companies.preorder_timeframes, ProductDetails.weight, (jewelry.title + ' ' + ISNULL(ProductDetails.Gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' +  ISNULL(ProductDetails.ProductDetail1,'')) AS 'mini_preview_text', (ISNULL(ProductDetails.Gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' +  ISNULL(ProductDetails.ProductDetail1,'')) AS 'variant', CAST(ProductDetails.price * tbl_carts.cart_qty as money) AS 'mini_line_price' FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN tbl_carts ON ProductDetails.ProductDetailID = tbl_carts.cart_detailId INNER JOIN TBL_Companies ON jewelry.brandname = TBL_Companies.name WHERE (tbl_carts." & var_db_field & " = ?) AND cart_save_for_later = 0 " & var_addons & " AND ProductDetails.active = 1 AND jewelry.active = 1 AND customorder='yes'"
objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10,var_cart_userid))
Set rsCart = objCmd.Execute()

While Not rsCart.eof
%>
<div style="float:left; clear: both">
	<img class="float-left mr-2 mb-1"  src="https://s3.amazonaws.com/bodyartforms-products/<%= rsCart.Fields.Item("picture").Value %>" alt="Product photo">
	<%= rsCart.Fields.Item("title").Value %>
</div>
<%
rsCart.MoveNext
Wend

DataConn.Close()
Set DataConn = Nothing
%>