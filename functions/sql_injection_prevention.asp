<%
	' protection to stop page from sql injection // for admin pages that do auto updating
	if column_title <> "" then
		if Len(column_title) > 30 then
			response.end
		end if
	end if
%>