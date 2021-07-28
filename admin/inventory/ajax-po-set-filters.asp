<%
if request.querystring("filter_status") <> "" then
	response.cookies("po-filter-status") = request.querystring("filter_status")
end if

if request.querystring("filter_active") <> "" then
	response.cookies("po-filter-active") = request.querystring("filter_active")
end if

if request.querystring("filter_qty") <> "" then
	response.cookies("po-filter-qty") = request.querystring("filter_qty")
end if


if request.querystring("filter_autoclave") <> "" then
	response.cookies("po-filter-autoclave") = request.querystring("filter_autoclave")
	else
	response.cookies("po-filter-autoclave") = ""
end if
%>