<%
active_search = "btn-purple"
active_shipping = "btn-purple"
active_billing = "btn-purple"
active_profile = "btn-purple"
active_orders = "btn-purple"
active_credits = "btn-purple"
active_wishlist = "btn-purple"
active_waiting = "btn-purple"

if Request.ServerVariables("URL") = "/account-searches.asp" then 
	active_search = "btn-info"
elseif Request.ServerVariables("URL") = "/account-shipping.asp" then 
	active_shipping = "btn-info"
elseif Request.ServerVariables("URL") = "/account-billing.asp" then 
	active_billing = "btn-info"
elseif Request.ServerVariables("URL") = "/account-profile.asp" then 
	active_profile = "btn-info"
elseif Request.ServerVariables("URL") = "/account.asp" then 
	active_orders = "btn-info"
elseif Request.ServerVariables("URL") = "/account-credits.asp" then 
	active_credits = "btn-info"
elseif Request.ServerVariables("URL") = "/wishlist.asp" then 
	active_credits = "btn-info"
elseif Request.ServerVariables("URL") = "/account-waiting-list.asp" then 
	active_waiting = "btn-info"
end if 
%>
<% if Not rsGetUser.EOF then %>
<% if rsGetUser("active") = True then %>
<ul class="nav nav-pills my-3" id="account-tabs">
		<li class="nav-item m-1">
				<a class="nav-link btn btn-sm <%= active_orders %>" href="account.asp">Orders</a>
		</li>
		
		<li class="nav-item m-1">
				<a class="nav-link btn btn-sm <%= active_billing %>" href="account-billing.asp">Credit Cards</a>
		</li>
		<li class="nav-item m-1">
				<a class="nav-link btn btn-sm <%= active_shipping %>" href="account-shipping.asp">Addresses</a>
		</li>
		<li class="nav-item m-1">
				<a class="nav-link btn btn-sm <%= active_credits %>" href="account-credits.asp">Credits</a>
		</li>
		<% if not rsNavWishlist.eof then %>
		<li class="nav-item m-1">
			<a class="nav-link btn btn-sm <%= active_wishlist %>" href="wishlist.asp">Wishlist</a>
		</li>
		<% end if %>
		<% if not rsNavWaitingList.eof then %>
		<li class="nav-item m-1">
			<a class="nav-link btn btn-sm <%= active_waiting %>" href="account-waiting-list.asp">Waiting List</a>
		</li>
		<% end if %>
		<% if not rsNavSavedSearches.eof then %>
		<li class="nav-item m-1">
				<a class="nav-link btn btn-sm <%= active_search %>" href="account-searches.asp">Saved Searches</a>
		</li>
		<% end if %>
		<li class="nav-item m-1">
			<a class="nav-link btn btn-sm <%= active_profile %>" href="account-profile.asp">Profile</a>
		</li>
		
</ul>
<% end if '===== only show navigation if account is active %>
<% end if %>