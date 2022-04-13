<% @LANGUAGE="VBSCRIPT" %>
<%
	page_title = "Wishlist"
	page_description = "Bodyartforms wishlist"
	page_keywords = ""
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->

<%
total_records = 0
sqlstring = ""

	if request.querystring("compact") = "yes" then
		session("wishlist-compact-mode") = "yes"
	end if
	if request.querystring("compact") = "no" then
		session("wishlist-compact-mode") = ""
	end if

	if session("wishlist_orderby") = "" then
		session("wishlist_orderby") = "ORDER BY ID DESC"
		session("wishlist_friendly_orderby") = "Newest first (Default)"
	end if

	if session("wishlist_jewelry") <> "" then
		sqlstring = sqlstring & " AND jewelry LIKE '%' + ? + '%'"
	end if
	if session("wishlist_gauge") <> "" then
		sqlstring = sqlstring & " AND Gauge = ?"
	end if 
	if session("wishlist_list") <> "" then
		sqlstring = sqlstring & " AND TBL_WishlistCategories.WishlistID = ?"
	end if 	
	if session("wishlist_brand") <> "" then
		sqlstring = sqlstring & " AND brandname LIKE '%' + ? + '%'"
	end if 
	if session("wishlist_keywords") <> "" then
		sqlstring = sqlstring & " AND (ISNULL(Gauge,'') + ' ' + ISNULL(Length,'') + ' ' + ISNULL(ProductDetail1,'') + ' ' + ISNULL(jewelry.title,'') + ' ' + ISNULL(brandname,'') + ' ' + ISNULL(jewelry,'')) LIKE '%' + ? + '%'"
	end if 	
	
'	response.write session("wishlist_keywords") & "<br/>" & sqlstring
	
'	response.write sqlstring & "<br/> " & session("wishlist_orderby") & "<br/>" & session("wishlist_list") & "<br/>" & session("wishlist_gauge")

' Pull the customer information from a cookie or userID in querystring
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT customer_ID FROM customers  WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
Set rsLoggedInUser = objCmd.Execute()

	var_user_id = 0
	var_user_status = ""
	If Not rsLoggedInUser.EOF then
		if int(CustID_Cookie) = int(rsLoggedInUser.Fields.Item("customer_ID").Value) then
			var_user_id = CustID_Cookie
			var_user_status = "own"
		'	response.write "Logged in user found"
		end if
	end if	

	if  var_user_id = 0 then ' if logged in user is not found
		if request.querystring("userID") <> "" then
			var_user_id = request.querystring("userID")
		else 
			var_user_id = 0
		end if
		var_user_status = "other"
	'	response.write "Guest user"
	end if

	If IsNumeric(var_user_id) Then
		' leave var_user_id as is
	Else
		' reset to 0, prevents asp errors from random entries in querystring
		var_user_id = 0
	End If

	if var_user_id <> 0 and request.querystring("userID") <> "" AND request.querystring("userID") <> var_user_id then ' user is logged in but viewing another persons wishlist
	'	response.write "Logged in but viewing someone elses wishlist"
		var_user_id = request.querystring("userID")
		var_user_status = "other"
	end if

'	response.write "VARIABLE USER ID " & var_user_id & ", COOKIE ID " & CustID_Cookie

' Pull the customer information from a cookie or userID in querystring
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM customers  WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,var_user_id))
Set rsGetUser = objCmd.Execute()

	If Not rsGetUser.EOF then
		wishlist_name = rsGetUser.Fields.Item("customer_first").Value
	else
		wishlist_name = ""
	end if
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.Prepared = true
	objCmd.CommandText = "SELECT  TOP (100) PERCENT wishlist.ID, wishlist.custID, wishlist.itemID, ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '') + ' ' + ISNULL(jewelry.title, '') AS title, jewelry.title AS title_sort, ProductDetails.price, jewelry.picture, jewelry.picture AS jewelry_thumb, jewelry.picture_400, largepic, jewelry.customorder, jewelry.type,  jewelry.brandname, wishlist.itemDetailID, ProductDetails.qty, customers.customer_first, customers.wishlist_nickname, customers.email, customers.customer_ID, jewelry.customorder, wishlist.desired, WishlistName, wishlist.priority, wishlist.waiting_list, wishlist.comments, wishlist.dateadded, jewelry.GroupDiscCategory, jewelry.jewelry, wishlist.WishlistID, CASE WHEN jewelry.active = 0 OR ProductDetails.active = 0 THEN 0 ELSE 1 END AS active, jewelry.ProductID, jewelry.SaleDiscount, wishlist.purchased, CASE WHEN SaleDiscount > 0 THEN ((price / 100) * (100 - SaleDiscount) * desired) ELSE (price * desired) END AS total_price, tbl_images.img_thumb, ProductDetails.img_id, CASE WHEN priority = 1 THEN 'Must have' WHEN priority = 2 THEN 'Love to have' WHEN priority = 3 THEN 'Like to have' WHEN priority = 4 THEN 'I am thinking about it' WHEN priority = 5 THEN 'Do not buy this for me' END AS wish_priority, TBL_Companies.ShowTextLogo, TBL_Companies.ProductLogo FROM TBL_WishlistCategories RIGHT OUTER JOIN ProductDetails RIGHT OUTER JOIN jewelry INNER JOIN TBL_Companies ON jewelry.brandname = TBL_Companies.name RIGHT OUTER JOIN customers INNER JOIN wishlist ON customers.customer_ID = wishlist.custID ON jewelry.ProductID = wishlist.itemID ON ProductDetails.ProductDetailID = wishlist.itemDetailID ON TBL_WishlistCategories.WishlistID = wishlist.WishlistID LEFT OUTER JOIN tbl_images ON ProductDetails.img_id = tbl_images.img_id WHERE (wishlist.custID = ?) AND title IS NOT NULL AND itemID IS NOT NULL " & sqlstring & " " & session("wishlist_orderby") & ""
		
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,var_user_id))
	
	if session("wishlist_jewelry") <> "" then
		objCmd.Parameters.Append(objCmd.CreateParameter("Jewelry",200,1,30,session("wishlist_jewelry")))
	end if
	If session("wishlist_gauge") <> "" then 
		objCmd.Parameters.Append(objCmd.CreateParameter("Gauge",200,1,10,session("wishlist_gauge")))
	end if 	
	If session("wishlist_list") <> "" then
		objCmd.Parameters.Append(objCmd.CreateParameter("list",3,1,10,session("wishlist_list")))
	end if 
	If session("wishlist_brand") <> "" then 
		objCmd.Parameters.Append(objCmd.CreateParameter("brand",200,1,50,session("wishlist_brand")))
	end if 	
	If session("wishlist_keywords") <> "" then 
		objCmd.Parameters.Append(objCmd.CreateParameter("keywords",200,1,50,session("wishlist_keywords")))
	end if 	
	
	set rsGetWishlist = Server.CreateObject("ADODB.Recordset")
	rsGetWishlist.CursorLocation = 3 'adUseClient
	rsGetWishlist.Open objCmd
	rsGetWishlist.PageSize = 25
	total_records = rsGetWishlist.RecordCount
	intPageCount = rsGetWishlist.PageCount


	
	
	
	' Variables for paging
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
	



	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM TBL_WishlistCategories WHERE Wishlist_CustomerID = ? ORDER BY WishlistName ASC"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
	Set rsGetCategories = objCmd.Execute()
	
	If session("wishlist_list") <> "" then
		wishlist_list_id = session("wishlist_list")
	else
		wishlist_list_id = 0
	end if
		
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT WishlistName FROM TBL_WishlistCategories WHERE Wishlist_CustomerID = ? AND WishlistID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
	objCmd.Parameters.Append(objCmd.CreateParameter("wishlist_id",3,1,10,wishlist_list_id))
	Set rsGetListName = objCmd.Execute()
	
		
%>
<input type="hidden" id="user-id" value="<%= var_user_id %>">
<div class="display-5 mb-1">
	<% 
	If Not rsGetWishlist.EOF then 
	if var_user_status = "own" then
	%>
		Your Wishlist (<%= total_records %> items)
		
	<% else %>
			<%= wishlist_name %>'s wishlist
	<% end if ' logged in as self
	else ' show if no wishlist is found %>
		Wishlist
	<%
	end if ' Not rsGetWishlist.EOF
	 %>
</div>
	<% if CustID_Cookie <> 0 then %>
		<% if var_user_status = "own" then %>
			<button class="btn btn-sm btn-purple" id="btn-manage-lists" type="button" data-toggle="modal" data-target="#ManageLists" data-dismiss="modal" href="#">Manage your lists</button>
			<button class="btn btn-sm btn-purple icon-wishlist-share">
				<i class="fa fa-share fa-lg mr-1"></i> Share List
			</button>
			<div class="alert alert-info p-2 my-2 share-box" style="display:none">
				Your link to share with family &amp; friends is below:
				<p>
				https://www.bodyartforms.com/wishlist.asp?userid=<%= var_user_id %>
				</p>
				<button class="btn btn-sm btn-info copy-link" data-clipboard-text="https://www.bodyartforms.com/wishlist.asp?userid=<%= var_user_id %>"><i class="fa fa-share fa-lg mr-2"></i>Copy link</button>
				<div class="link-copied alert alert-success p-2 my-2" style="display:none">Link as been copied to your clipboard</div>
			</div>
		
		<% end if %>
	<% end if %>


<!--#include virtual="/wishlist/inc-wishlist-filters.asp" --> 


<% If rsGetWishlist.EOF Then %>
	<div class="alert alert-primary my-3">
		No wishlist items found
		<% if session("wishlist_keywords") <> "" then %>
			for keywords <%= session("wishlist_keywords") %>
		<% end if %>
	</div>
<% End If %>
<div class="mt-5"></div>
<% If Not rsGetWishlist.EOF then %>
<!--#include virtual="/wishlist/inc-wishlist-paging.asp" -->

<div class="d-flex flex-row flex-wrap">
<% 
rsGetWishlist.AbsolutePage = intPage '======== PAGING
For intRecord = 1 To rsGetWishlist.PageSize 

	if rsGetWishlist.Fields.Item("total_price").Value <> "" then
		price = rsGetWishlist.Fields.Item("total_price").Value
	end if
%>


<div class="col-6 col-md-4 col-lg-4 col-xl-3 col-break1600-2 my-3 px-0 px-md-2" id="block-<%= rsGetWishlist.Fields.Item("ID").Value %>">	
		<div class="container-fluid p-0">
<div class="mx-1">
	<a class="mb-2 d-block" href="/productdetails.asp?ProductID=<%=(rsGetWishlist.Fields.Item("ProductID").Value)%>">
		<div class="position-relative">
		<% if rsGetWishlist.Fields.Item("img_id").Value <> 0 then 
		if instr(rsGetWishlist.Fields.Item("img_thumb").Value,"thumbnail") > 0 then
		%>

			<img class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetWishlist.Fields.Item("img_thumb").Value %>" data-img-name="<%= rsGetWishlist.Fields.Item("largepic").Value %>" alt="Thumbnail">
		<% else %>
			<img class="img-fluid" style="width:400px;height:auto" src="http://bodyartforms-products.bodyartforms.com/<%= rsGetWishlist.Fields.Item("img_thumb").Value %>" data-img-name="<%= rsGetWishlist.Fields.Item("largepic").Value %>" alt="Product Photo">
		<% 
		end if ' if the word thumbnail is detected in image name, then it's safe to assume there is an image in the 400 bucket. If not, default to the basic thumbnail.
		else %>
			<img class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetWishlist.Fields.Item("picture_400").Value %>" data-img-name="<%= rsGetWishlist.Fields.Item("largepic").Value %>" alt="Thumbnail">
		<% end if %>
		<% if rsGetWishlist.Fields.Item("ShowTextLogo").Value = "Y" then 
		%>
		<div class="brand-info position-absolute w-50 badge badge-light rounded-0" style=" bottom:5px;right: 5px;overflow-wrap: break-word;">
			<img class="img-fluid" src="/images/<%= rsGetWishlist.Fields.Item("ProductLogo").Value %>" alt="logo" />
		</div>
		<% end if %>
		</div><!-- relative position-->
	</a>
	<% if var_user_status = "own" then %>
		<button class="btn btn-sm btn-outline-secondary btn-block mb-1 btn-open-update-item" data-id="<%= rsGetWishlist.Fields.Item("ID").Value %>" data-toggle="modal" data-target="#updateWishItem"
			data-dismiss="modal">Update / Delete</button>
	<% end if %>
	<% if rsGetWishlist.Fields.Item("active").Value = 1 then %>

		<% if (rsGetWishlist.Fields.Item("qty").Value) > "0" then %>
		
			<button class="btn btn-purple btn-block m-0 mb-1 add-cart add-cart-<%= rsGetWishlist.Fields.Item("ID").Value %>" data-id="<%= rsGetWishlist.Fields.Item("ID").Value %>" data-detailid="<%= rsGetWishlist.Fields.Item("itemDetailID").Value %>" data-productid="<%= rsGetWishlist.Fields.Item("productid").Value %>"><i class="fa fa-shopping-cart fa-lg mr-2"></i><span class="mr-2"><%= FormatCurrency(price, -1, -2, -2, -2) %></span><span class="d-block d-md-inline">Add <span id="desired-<%= rsGetWishlist.Fields.Item("ID").Value %>"><%= rsGetWishlist.Fields.Item("desired").Value %></span> to cart</span></button>
			<div class="msg-add-cart-<%= rsGetWishlist.Fields.Item("ID").Value %>"></div>

		<% else ' if item is out of stock 
		
		' Don't show out of stock button if only the main product (with no item) was added to the list
		if rsGetWishlist.Fields.Item("itemDetailID").Value <> 0 AND rsGetWishlist.Fields.Item("itemDetailID").Value <> "" then
		%> 
			<div class="alert alert-secondary p-1 m-0 mb-1 text-center">OUT OF STOCK</div>
			<% else %>
			<a class="btn btn-purple btn-block m-0 mb-1" href="/productdetails.asp?ProductID=<%=(rsGetWishlist.Fields.Item("ProductID").Value)%>">View product</a>
		<% end if %>
			
			<% if var_user_status = "own" then %>
			<% if rsGetWishlist.Fields.Item("waiting_list").Value = 0 then %>
			<button class="btn btn-sm btn-outline-secondary btn-block add-waiting" type="button" data-detailid="<%= rsGetWishlist.Fields.Item("itemDetailID").Value %>" data-wishlistid="<%= rsGetWishlist.Fields.Item("ID").Value %>"><i class="fa fa-email"></i> Notify me when re-stocked</button> 
			<% else ' if already on waiting list %>
				<i class="fa fa-check btn btn-sm btn-outline-success"></i> 
				You're on the waiting list
			<% end if 'if on waiting list 
			end if ' if logged in as own user %>
			
		<% end if ' if item is out/in stock %> 
		
	<% else ' if active = 0 %>
		<div class="alert alert-danger text-center p-1 m-0 mb-1">Discontinued item</div>
	<% end if ' if active = 1 %>


		<div class="font-weight-bold small">
			<%=(rsGetWishlist.Fields.Item("title").Value)%>
		</div>
		<% if rsGetWishlist.Fields.Item("comments").Value <> "" then %>
		<div class="small">
				<span class="font-weight-bold mr-2">Specs:</span><span id="comments-<%= rsGetWishlist.Fields.Item("ID").Value %>"><%= Server.HTMLEncode(rsGetWishlist.Fields.Item("comments").Value & "") %></span>
		</div>
		<% end if %>
		<div class="small">
				<span class="font-weight-bold mr-2">Priority:</span><span id="priority-<%= rsGetWishlist.Fields.Item("ID").Value %>"><%= rsGetWishlist.Fields.Item("wish_priority").Value %></span>
		</div>
		<% if var_user_status = "own" then 
		if rsGetWishlist.Fields.Item("WishlistName").Value <> "" then %>
		<div class="small">
				<span class="font-weight-bold mr-2">In List:</span><span id="category-<%= rsGetWishlist.Fields.Item("ID").Value %>"><%= rsGetWishlist.Fields.Item("WishlistName").Value %></span>
		</div>
		<% end if 
		end if %>
		<div class="small text-secondary">
				Added to list <%=(rsGetWishlist.Fields.Item("dateadded").Value)%>
			</div>

	<% if rsGetWishlist.Fields.Item("purchased").Value = 1 then ' only show add to cart if item has not been purchased %>
		<div class="alert alert-primary py-0 px-2 my-1">THIS ITEM HAS BEEN PURCHASED</div>
	<% end if %>

</div><!-- margin padding -->
	</div><!-- container fluid -->
</div><!--flex columns-->
<% 
rsGetWishlist.MoveNext()
If rsGetWishlist.EOF Then Exit For  ' ====== PAGING
Next ' ====== PAGING
%>
</div><!--flex-->


<% if var_user_status = "own" then %>
        <!-- Update wishlist item modal -->
        <div class="modal fade" id="updateWishItem" tabindex="-1" role="dialog" aria-labelledby="updateWishItemLabel"
                aria-hidden="true">
                <div class="modal-dialog" role="document">
                        <div class="modal-content">
                                <div class="modal-header">
                                        <h5 class="modal-title" id="updateWishItemLabel">Update Wishlist Item</h5>
                                        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                                                <span aria-hidden="true">&times;</span>
                                        </button>
                                </div>
                                <div class="modal-body">
									<input type="hidden" id="update-id">
									<div id="loader-update-item"></div>	
									<div id="spinner-update-item" style="display:none"><i class="fa fa-spinner fa-2x fa-spin"></i></div>
									<div id="message-update-item"></div>	
								</div>
								<div class="modal-footer">
										<button type="button" class="btn btn-danger" id="delete-wishlist-item">Delete Item</button>
										<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
										<button type="submit" class="btn btn-purple" id="update-wishlist-item">Save Changes</button>
								</div>
                        </div>
                </div>
        </div>
        <!-- Manage lists Modal -->
        <div class="modal fade" id="ManageLists" tabindex="-1" role="dialog" aria-labelledby="ManageListsLabel"
                aria-hidden="true">
                <div class="modal-dialog" role="document">
                        <div class="modal-content">
                                <div class="modal-header">
                                        <h5 class="modal-title" id="ManageListsLabel">Manage Your Lists</h5>
                                        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                                                <span aria-hidden="true">&times;</span>
                                        </button>
                                </div>
                                <div class="modal-body">
									<div class="form-inline">
										<input class="form-control form-control-sm w-auto d-inline" type="text" maxlength="50" name="add-category" id="add-category" placeholder="Enter a new list name" /><button class="btn btn-sm btn-outline-success ml-2 btn-add-category" type="button">Create New List</button>
									</div>	
								
									<div class="category-spinner" style="display:none"><i class="fa fa-spinner fa-2x fa-spin"></i></div>
									<div class="category-message"></div><!-- jquery message div -->
									<div class="manage-categories"></div><!-- jquery loader div -->
                                </div>
                        </div>
                </div>
        </div>
<% end if ' var_user_status = "own" %>



<!--#include virtual="/wishlist/inc-wishlist-paging.asp" -->



<% end if ' If Not rsGetWishlist.EOf %>                  


<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript" src="/js/clipboard.js"></script>
<script type="text/javascript" src="/js-pages/wishlist.min.js?v=051519"></script>