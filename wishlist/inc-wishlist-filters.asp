<%
	if session("wishlist_jewelry") <> "" then
		selected_jewelry = session("wishlist_jewelry")
	end if
	
	if session("wishlist_brand") <> "" then
		selected_brand = session("wishlist_brand")
	end if
	
	if session("wishlist_material") <> "" then
		selected_material = session("wishlist_material")
	end if
%>
<form class="mt-4" id="wishlist-filters" action="?userID=<%= request.querystring("userID") %>" method="get">
<div class="mb-1" id="wishlist-search">
	<input class="form-control form-field-auto d-md-inline" type="text" name="wishlist-search" placeholder="Wishlist keyword search">
	<div class="mt-2 d-md-none"></div>
	<input class="btn btn-sm btn-purple" type="submit" id="btn-submit-filters" value="Search">
	<button class="btn btn-sm btn-purple d-md-none link-expand-filters" type="button">
	Show more filters <span class="wishlist-filter-down"><i class="fa fa-chevron-down"></i></span><span class="wishlist-filter-up" style="display:none"><i class="fa fa-chevron-up"></i></span>
	</button>
	<div class="mb-3"></div>
</div>
<% if var_user_status = "own" and NOT rsGetCategories.eof then %>
	<select class="my-1 form-control form-field-auto d-md-inline" name="wishlist-list" id="wishlist-list">
	<% if NOT rsGetListName.eof then %>
	<option value="<%= session("wishlist_list") %>" selected="selected"><%=(rsGetListName.Fields.Item("WishlistName").Value)%></option>
	<% end if
	%>
	<option value="">Filter by list...</option>
	<option value="">No filter (show all)</option>
	<%
	If Not rsGetCategories.EOF Then
	While NOT rsGetCategories.EOF 
		if rsGetCategories.Fields.Item("Wishlist_CustomerID").Value <> 0 then
	%>
	<option value="<%=(rsGetCategories.Fields.Item("WishlistID").Value)%>"><%=(rsGetCategories.Fields.Item("WishlistName").Value)%></option>
	<% 
		end if ' Wishlist_CustomerID <> 0
	rsGetCategories.MoveNext()
	Wend
	End If ' end Not rsGetCategories.EOF 
	%>
	</select>
<% end if %>
<div class="d-md-inline" style="display:none" id="expand-filters">
	
<select class="my-1 form-control form-field-auto d-md-inline" name="wishlist-sort" id="wishlist-sort">
<% if session("wishlist_orderby") <> "" then %>
<option value="<%= session("wishlist_orderby") %>" selected="selected"><%= session("wishlist_friendly_orderby") %></option>
<% end if %>

<option value="">SORT BY:</option>
<option value="dateadded DESC">Default</option>
<option value="price ASC">Price (low to high)</option>
<option value="price DESC">Price (high to low)</option>
<option value="purchased">Purchased items first</option>
<option value="limited">Limited items first</option>
<option value="priority ASC">Priority (high to low)</option>
<option value="priority DESC">Priority (low to high)</option>
<option value="dateadded ASC">Date added (old to new)</option>
<option value="dateadded DESC">Date added (new to old)</option>
<option value="title ASC">Title (A to Z)</option>
</select>

<select class="my-1 form-control form-field-auto d-md-inline" name="wishlist-jewelry" id="wishlist-jewelry">            
	<!--#include virtual="/template/inc_jewelry_select.asp" -->
</select>

<select class="my-1 form-control form-field-auto d-md-inline" name="wishlist-gauge" id="wishlist-gauge">
<% if session("wishlist_gauge") <> "" then %>
<option value="<%= Server.HTMLEncode(session("wishlist_gauge")) %>" selected="selected"><%= Server.HTMLEncode(session("wishlist_gauge")) %></option>
<% end if %>          
<option value="">Gauge...</option>
<option value="">No filter (show all)</option>    
<option value="20g">20g</option>
<option value="18g">18g</option>
<option value="16g">16g</option>
<option value="14g">14g</option>
<option value="12g">12g</option>
<option value="10g">10g</option>
<option value="8g">8g</option>
<option value="6g">6g</option>
<option value="4g">4g</option>
<option value="1g">1g</option>
<option value="2g">2g</option>
<option value="0g">0g</option>
<option value="00g">00g</option>
<option value="00g/9mm">00g/9mm</option>
<option value="00g/9.5mm">00g/9.5mm</option>
<option value="00g/10mm">00g/10mm</option>
<option value="7/16&quot;">7/16&quot;</option>
<option value="12mm">12mm</option>
<option value="1/2&quot;">1/2&quot;</option>
<option value="9/16&quot;">9/16&quot;</option>
<option value="15mm">15mm</option>
<option value="5/8&quot;">5/8&quot;</option>
<option value="11/16&quot;">11/16&quot;</option>
<option value="18mm">18mm</option>
<option value="3/4&quot;">3/4&quot;</option>
<option value="19mm">19mm</option>
<option value="13/16&quot;">13/16&quot;</option>
<option value="7/8&quot;">7/8&quot;</option>
<option value="15/16&quot;">15/16&quot;</option>
<option value="25mm">25mm</option>
<option value="1&quot;">1&quot;</option>
<option value="1-1/16&quot;">1-1/16&quot;</option>
<option value="1-1/8&quot;">1-1/8&quot;</option>
<option value="1-3/16&quot;">1-3/16&quot;</option>
<option value="1-1/4&quot;">1-1/4&quot;</option>
<option value="1-5/16&quot;">1-5/16&quot;</option>
<option value="1-3/8&quot;">1-3/8&quot;</option>
<option value="1-7/16&quot;">1-7/16&quot;</option>
<option value="1-1/2&quot;">1-1/2&quot;</option>
<option value="1-9/16&quot;">1-9/16&quot;</option>
<option value="1-5/8&quot;">1-5/8&quot;</option>
<option value="1-11/16&quot;">1-11/16&quot;</option>
<option value="1-3/4&quot;">1-3/4&quot;</option>
<option value="1-7/8&quot;">1-7/8&quot;</option>
<option value="2&quot;">2&quot;</option>
<option value="2-1/16&quot;">2-1/16&quot;</option>
<option value="2-1/8&quot;">2-1/8&quot;</option>
<option value="2-3/16&quot;">2-3/16&quot;</option>
<option value="2-1/4&quot;">2-1/4&quot;</option>
<option value="2-5/16&quot;">2-5/16&quot;</option>
<option value="2-3/8&quot;">2-3/8&quot;</option>
<option value="2-7/16&quot;">2-7/16&quot;</option>
<option value="2-1/2&quot;">2-1/2&quot;</option>
<option value="2-9/16&quot;">2-9/16&quot;</option>
<option value="2-5/8&quot;">2-5/8&quot;</option>
<option value="2-3/4&quot;">2-3/4&quot;</option>
<option value="2-7/8&quot;">2-7/8&quot;</option>
<option value="3&quot;">3&quot;</option>
</select>

<select class="my-1 form-control form-field-auto d-md-inline" name="wishlist-brand" id="wishlist-brand" >
	<!--#include virtual="/template/inc_brand_select.asp" -->
</select>

<div class="d-block">
<button class="btn btn-sm btn-purple clear-filters" type="button">
		Clear all filters
	</button>
</div>
</div>
</form> 