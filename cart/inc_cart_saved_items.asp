<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="cart/generate_guest_id.asp"-->
<% 
if CustID_Cookie <> 0 then

	' RETRIEVE SAVE FOR LATER CONTENTS ---------------------------------------------
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT tbl_carts.cart_id, tbl_carts.cart_detailID, tbl_carts.cart_preorderNotes, tbl_carts.cart_custId, tbl_carts.cart_qty, tbl_carts.cart_wishlistid, jewelry.customorder, jewelry.SaleDiscount, jewelry.secret_sale, jewelry.ProductID, jewelry.active as p_active, jewelry.picture, ProductDetails.active as d_active, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.price, ProductDetails.qty, ProductDetails.ProductDetail1, ProductDetails.ProductDetailID, (jewelry.title + ' ' + ProductDetails.Gauge + ' ' + ProductDetails.Length + ' ' +  ProductDetails.ProductDetail1) AS 'title' FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN tbl_carts ON ProductDetails.ProductDetailID = tbl_carts.cart_detailId WHERE (tbl_carts." & var_db_field & " = ?) AND cart_save_for_later = 1"
	objCmd.Parameters.Append(objCmd.CreateParameter("cart_custID",3,1,10,var_cart_userid))
	
	set rs_getSaveForLater = Server.CreateObject("ADODB.Recordset")
	rs_getSaveForLater.CursorLocation = 3 'adUseClient
	rs_getSaveForLater.Open objCmd
	rs_getSaveForLater.PageSize = 5 ' not using (possibly needed for pagination)
	intPageCount = rs_getSaveForLater.PageCount ' not using (possibly needed for pagination)

Select Case Request("Action")
	case "<<"
		intpage = 1
	case "<"
		intpage = Request("intpage")-1
		if intpage < 1 then intpage = 1
	case ">"
		intpage = Request("intpage")+1
		if intpage > intPageCount then intpage = IntPageCount
	Case ">>"
		intpage = intPageCount
	case else
		intpage = 1
end select


if NOT rs_getSaveForLater.eof then 
%>
<div class="card">
	<div class="card-header">
			<h5>Your saved for later items</h5>
	</div>
<div class="card-body p-2">
	<!--#include virtual="cart/inc_save_for_later_paging.asp"-->
	<div class="container-fluid">
<%
rs_getSaveForLater.AbsolutePage = intPage '======== PAGING
For intRecord = 1 To rs_getSaveForLater.PageSize 
%>
<div class="row my-2 detailid_<%= rs_getSaveForLater.Fields.Item("cart_id").Value %>">
	<div class="col-auto m-0 p-0 mr-2">		  
	
	<a href="productdetails.asp?ProductID=<%=(rs_getSaveForLater.Fields.Item("ProductID").Value)%>"><img src="https://bodyartforms-products.bodyartforms.com/<%=(rs_getSaveForLater.Fields.Item("picture").Value)%>" alt="Product photo"></a>
	</div>	
	<div class="col p-0">	 
		<div class="small">				
			<%= rs_getSaveForLater.Fields.Item("title").Value %>
			
			<% if rs_getSaveForLater.Fields.Item("cart_preorderNotes").Value <> "" then %>	  
			<% if rs_getSaveForLater.Fields.Item("ProductID").Value <> 2424 then ' if item is not a gift certificate %>
			<div>
				Your specs: <%=(rs_getSaveForLater.Fields.Item("cart_preorderNotes").Value)%>
			</div>
			<% else ' show gift certificate information 
				certificate_array =split(rs_getSaveForLater.Fields.Item("cart_preorderNotes").Value,"{}")				
			%>
			<div>
			Gift certificate information:
			Recipient's name: <%= certificate_array(3) %>
			Recipient's e-mail: <%= certificate_array(0) %>
			Your name: <%= certificate_array(1) %>
			Your message: <%= certificate_array(2) %>
		</div>
			<%	end if ' detect whether preorder or gift cert %>
		<% end if %>
	</div>
				
				<div class="font-weight-bold">
						<%= exchange_symbol %><%= FormatNumber(rs_getSaveForLater.Fields.Item("price").Value * exchange_rate, -1, -2, -2, -2) %>
						<% 	if (rs_getSaveForLater.Fields.Item("SaleDiscount").Value > 0 AND rs_getSaveForLater.Fields.Item("secret_sale").Value = 0) OR (rs_getSaveForLater.Fields.Item("secret_sale").Value = 1 AND session("secret_sale") = "yes") then %>
						<span class="badge badge-danger">ON SALE <%= rs_getSaveForLater.Fields.Item("SaleDiscount").Value %>% OFF</span>
					  <% end if %>
					</div>
				

		
			<% var_in_stock = ""
			if rs_getSaveForLater.Fields.Item("cart_qty").Value <= rs_getSaveForLater.Fields.Item("qty").Value then 
				var_in_stock = "yes"
			end if %>
			<div class="mt-2">
			<button class="btn btn-sm btn-outline-danger action-remove mr-2" data-detailid="<%= rs_getSaveForLater.Fields.Item("cart_id").Value %>" data-specs="<%= rs_getSaveForLater.Fields.Item("cart_id").Value %>" type="button"><i class="fa fa-times"></i></button>

				<% if var_in_stock = "yes" and rs_getSaveForLater.Fields.Item("d_active").Value = 1 and rs_getSaveForLater.Fields.Item("p_active").Value = 1 then
				 %>
				
					<a class="btn btn-sm btn-outline-secondary" href="?remove_save=<%= rs_getSaveForLater.Fields.Item("cart_id").Value %>" title="Move to cart">Move to cart</a>
				
			<% else 
				if rs_getSaveForLater.Fields.Item("d_active").Value = 0 or rs_getSaveForLater.Fields.Item("p_active").Value = 0 then
					savelater_status = " (Discontinued)"
				end if
			%>
			
				<strong>OUT OF STOCK <%= savelater_status %></strong>
			
			<% end if %>
		</div>
	</div><!-- col -->
</div><!-- row -->
<hr>
<%
rs_getSaveForLater.MoveNext()
If rs_getSaveForLater.EOF Then Exit For  ' ====== PAGING
Next ' ====== PAGING
%>
</div><!-- container-fluid -->
	<!--#include virtual="cart/inc_save_for_later_paging.asp"-->
</div><!-- card body -->
</div><!-- card -->
<% 
end if 'if NOT rs_getSaveForLater.eof 
end if ' CustID_Cookie <> ""

DataConn.Close()
Set DataConn = Nothing

%>