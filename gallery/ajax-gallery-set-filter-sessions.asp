<%
' Clear all filters
If Request.form("clear") = "yes" then 
	session("gallery_jewelry") = ""
	session("gallery_gauge") = ""
	session("gallery_list") = ""
	session("gallery_brand") = ""
	session("gallery_orderby") = ""
	session("gallery_custid") = ""
	session("gallery_productid") = ""
end if

' Set search keyword
If Request.form("gallery-search") <> "" then 
	session("gallery_keywords") = request.form("gallery-search")
else
	session("gallery_keywords") = ""
end if 

' Set jewelry filter
If Request.form("gallery-jewelry") <> "" then 
	session("gallery_jewelry") = request.form("gallery-jewelry")
else
	session("gallery_jewelry") = ""
end if 

' Set gauge filter
If Request.form("gallery-gauge") <> "" then 
	session("gallery_gauge") = request.form("gallery-gauge")
else
	session("gallery_gauge") = ""
end if 

' Set list filter
If Request.form("gallery-material") <> "" then 
	session("gallery_material") = request.form("gallery-material")
else
	session("gallery_material") = ""
end if 

' Set brand filter
If request.form("gallery-brand") <> "" then 
	session("gallery_brand") = request.form("gallery-brand")
else
	session("gallery_brand") = ""
end if 

' Set product id
If request.form("gallery-productid") <> "" then 
	session("gallery_productid") = request.form("gallery-productid")
else
	session("gallery_productid") = ""
end if 


' Set order by
If Request.form("gallery-sort") = "jewelry ASC, title ASC" then
	session("gallery_orderby") = "ORDER BY jewelry ASC, title ASC"
ElseIf Request.form("gallery-sort") = "price ASC" then
	session("gallery_orderby") = "ORDER BY price ASC"
	session("gallery_friendly_orderby") = "Least expensive"
ElseIf Request.form("gallery-sort") = "price DESC" then
	session("gallery_orderby") = "ORDER BY price DESC"
	session("gallery_friendly_orderby") = "Most expensive"
ElseIf Request.form("gallery-sort") = "priority ASC" then
	session("gallery_orderby") = "ORDER BY priority ASC"
	session("gallery_friendly_orderby") = "Least important first"
ElseIf Request.form("gallery-sort") = "priority DESC" then
	session("gallery_orderby") = "ORDER BY priority DESC"
	session("gallery_friendly_orderby") = "Most important first"
ElseIf Request.form("gallery-sort") = "dateadded ASC" then
	session("gallery_orderby") = "ORDER BY dateadded ASC"
	session("gallery_friendly_orderby") = "Oldest first"
ElseIf Request.form("gallery-sort") = "dateadded DESC" then
	session("gallery_orderby") = "ORDER BY dateadded DESC"
	session("gallery_friendly_orderby") = "Newest first"
ElseIf Request.form("gallery-sort") = "title ASC" then
	session("gallery_orderby") = "ORDER BY title_sort ASC"
	session("gallery_friendly_orderby") = "Title (A to Z)"
ElseIf Request.form("gallery-sort") = "purchased" then
	session("gallery_orderby") = "ORDER BY purchased DESC"
	session("gallery_friendly_orderby") = "Purchased items first"
ElseIf Request.form("gallery-sort") = "limited" then
	session("gallery_orderby") = "ORDER BY CASE WHEN type = 'limited' THEN '1' WHEN type = 'onetime' THEN '2' WHEN type = 'clearance' or type = 'discontinued' THEN '3' ELSE '4' END ASC"
	session("gallery_friendly_orderby") = "Limited items first"
End If  
%>