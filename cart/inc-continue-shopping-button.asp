<%
' Set page last viewed for continue shopping button. Only write if user is on products or productdetails page. Was using javascript local storage but it throws errors in private browsing mode.
	if Request.ServerVariables("URL") = "/products/ajax-products-display.asp" then
		var_page_cont_name = "/products.asp"
	else
		var_page_cont_name = "/productdetails.asp"
	end if
	
	' Decided to make it track ONLY the main products page 
		session("continue_shopping_link") = "/products.asp?" & Request.ServerVariables("QUERY_STRING")
%>