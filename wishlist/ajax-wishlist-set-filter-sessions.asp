<%
' Clear all filters
If Request.form("clear") = "yes" then 
	session("wishlist_jewelry") = ""
	session("wishlist_gauge") = ""
	session("wishlist_list") = ""
	session("wishlist_brand") = ""
	session("wishlist_orderby") = ""
end if

' Set search keyword
If Request.form("wishlist-search") <> "" then 
	session("wishlist_keywords") = request.form("wishlist-search")
else
	session("wishlist_keywords") = ""
end if 

' Set jewelry filter
If Request.form("wishlist-jewelry") <> "" then 
	session("wishlist_jewelry") = request.form("wishlist-jewelry")
else
	session("wishlist_jewelry") = ""
end if 

' Set gauge filter
If Request.form("wishlist-gauge") <> "" then 
	session("wishlist_gauge") = request.form("wishlist-gauge")
else
	session("wishlist_gauge") = ""
end if 

' Set list filter
If Request.form("wishlist-list") <> "" then 
	session("wishlist_list") = request.form("wishlist-list")
else
	session("wishlist_list") = ""
end if 

' Set brand filter
If request.form("wishlist-brand") <> "" then 
	session("wishlist_brand") = request.form("wishlist-brand")
else
	session("wishlist_brand") = ""
end if 


' Set order by
If Request.form("wishlist-sort") = "jewelry ASC, title ASC" then
	session("wishlist_orderby") = "ORDER BY jewelry ASC, title ASC"
ElseIf Request.form("wishlist-sort") = "price ASC" then
	session("wishlist_orderby") = "ORDER BY price ASC"
	session("wishlist_friendly_orderby") = "Least expensive"
ElseIf Request.form("wishlist-sort") = "price DESC" then
	session("wishlist_orderby") = "ORDER BY price DESC"
	session("wishlist_friendly_orderby") = "Most expensive"
ElseIf Request.form("wishlist-sort") = "priority ASC" then
	session("wishlist_orderby") = "ORDER BY priority ASC"
	session("wishlist_friendly_orderby") = "Most important"
ElseIf Request.form("wishlist-sort") = "priority DESC" then
	session("wishlist_orderby") = "ORDER BY priority DESC"
	session("wishlist_friendly_orderby") = "Least important"
ElseIf Request.form("wishlist-sort") = "dateadded ASC" then
	session("wishlist_orderby") = "ORDER BY dateadded ASC"
	session("wishlist_friendly_orderby") = "Oldest first"
ElseIf Request.form("wishlist-sort") = "dateadded DESC" then
	session("wishlist_orderby") = "ORDER BY dateadded DESC"
	session("wishlist_friendly_orderby") = "Newest first"
ElseIf Request.form("wishlist-sort") = "title ASC" then
	session("wishlist_orderby") = "ORDER BY title_sort ASC"
	session("wishlist_friendly_orderby") = "Title (A to Z)"
ElseIf Request.form("wishlist-sort") = "purchased" then
	session("wishlist_orderby") = "ORDER BY purchased DESC"
	session("wishlist_friendly_orderby") = "Purchased items first"
ElseIf Request.form("wishlist-sort") = "limited" then
	session("wishlist_orderby") = "ORDER BY CASE WHEN jewelry.type = 'limited' THEN '1' WHEN jewelry.type = 'onetime' THEN '2' WHEN jewelry.type = 'clearance' or jewelry.type = 'discontinued' THEN '3' ELSE '4' END ASC"
	session("wishlist_friendly_orderby") = "Limited items first"
End If  
%>