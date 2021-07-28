<!--#include virtual="/Connections/sql_connection.asp" -->
<%
if request.form("wishlist_id") <> "" then 

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT  TOP (100) PERCENT wishlist.ID, wishlist.custID, wishlist.itemID, ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '') + ' ' + ISNULL(jewelry.title, '') AS title, jewelry.title AS title_sort, ProductDetails.price, jewelry.picture, jewelry.picture AS jewelry_thumb, jewelry.customorder, jewelry.type,  jewelry.brandname, wishlist.itemDetailID, ProductDetails.qty, customers.customer_first, customers.wishlist_nickname, customers.email, customers.customer_ID, jewelry.customorder, wishlist.desired, WishlistName, wishlist.priority, wishlist.waiting_list, wishlist.comments, wishlist.dateadded, jewelry.GroupDiscCategory, jewelry.jewelry, wishlist.WishlistID, TBL_WishlistCategories.WishlistID AS 'category_id', CASE WHEN jewelry.active = 0 OR ProductDetails.active = 0 THEN 0 ELSE 1 END AS active, jewelry.ProductID, jewelry.OnSale, jewelry.SaleDiscount, wishlist.purchased, CASE WHEN OnSale = 'Y' THEN ((price / 100) * (100 - SaleDiscount) * desired) ELSE (price * desired) END AS total_price, tbl_images.img_thumb, ProductDetails.img_id, CASE WHEN priority = 1 THEN 'Must have' WHEN priority = 2 THEN 'Love to have' WHEN priority = 3 THEN 'Like to have' WHEN priority = 4 THEN 'I am thinking about it' WHEN priority = 5 THEN 'Do not buy this for me' END AS priority_friendly FROM tbl_images RIGHT OUTER JOIN TBL_WishlistCategories RIGHT OUTER JOIN jewelry RIGHT OUTER JOIN customers INNER JOIN wishlist ON customers.customer_ID = wishlist.custID ON jewelry.ProductID = wishlist.itemID LEFT OUTER JOIN ProductDetails ON wishlist.itemDetailID = ProductDetails.ProductDetailID ON TBL_WishlistCategories.WishlistID = wishlist.WishlistID ON tbl_images.img_id = ProductDetails.img_id WHERE wishlist.custID = ? AND wishlist.ID = ?"
    
    objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
    objCmd.Parameters.Append(objCmd.CreateParameter("wishlist_id",3,1,10,request.form("wishlist_id")))
    set rsGetWishlist = objCmd.Execute()


	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM TBL_WishlistCategories WHERE Wishlist_CustomerID = ? ORDER BY WishlistName ASC"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
	Set rsGetCategory = objCmd.Execute()
		
%>

<form id="frm-update-item">
<% if NOT rsGetWishlist.EOF then %>
<% ' Only display if lists are found
if NOT rsGetCategory.EOF then %>
<div class="form-group">
    <label for="category" id="category-select">Select List</label>
    <select class="form-control" name="category" id="category-select">
        <option value="0" selected >Select list...</option>
        <optgroup label="-------------------------"></optgroup>
        <%        While NOT rsGetCategory.EOF        %>
        <option value="<%= rsGetCategory.Fields.Item("WishlistID").Value %>" <%
         if rsGetWishlist.Fields.Item("category_id").Value = rsGetCategory.Fields.Item("WishlistID").Value then %>selected<% end if %>><%= rsGetCategory.Fields.Item("WishlistName").Value %></option>
        <% 
        rsGetCategory.MoveNext()
        Wend
        %>
    </select>
</div>
<% end if ' NOT rsGetCategory.EOF  %>

<div class="form-group">	
    <label for="desired">Quantity</label>
    <input class="form-control" name="desired" id="desired" maxlength="4" type="tel" class="qty" value="<%= rsGetWishlist.Fields.Item("desired").Value %>"> 
</div> 

<div class="form-group">
    <label for="priority">Priority</label>
    <select class="form-control" name="priority" id="priority">
        <option value="<%= rsGetWishlist.Fields.Item("priority").Value %>" selected>
            <%= rsGetWishlist.Fields.Item("priority_friendly").Value %>
        </option>
        <option value="1">1 - Must have</option>
        <option value="2">2 - Love to have</option>
        <option value="3">3 - Like to have</option>
        <option value="4">4 - I'm thinking about it</option>
        <option value="5">5 - Don't buy this for me</option>
    </select>    
</div>

<div class="form-group">
    <label for="comments">Comments or Pre-Order specs</label> 
    <input class="form-control" name="comments" id="comments" type="text" value="<%= rsGetWishlist.Fields.Item("comments").Value %>">
</div>


<% end if 'NOT rsGetWishlist.EOF %>
</form>
    

<% end if 'request.form("wishlist_id") <> "" %>