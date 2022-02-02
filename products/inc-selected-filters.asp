		<% if request.querystring <> "" then %>
			<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href="" data-filter="all">
			<i class="fa fa-times"></i>
			Remove all filters
			</a>
		<% end if %>
		<% if request.querystring("filter-stock") <> "all" and (request.querystring("gauge") <> "" or request.querystring("length") <> "" or request.querystring("price") <> "" or request.querystring("colors") <> "") then %>
			<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href=""  data-filter="stock">
			<i class="fa fa-times"></i>
			Showing in stock only
			</a>
		<% end if %>
		<% if request.querystring("filter-stock") = "all" then %>
			<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href=""  data-filter="stock-all">
			<i class="fa fa-times"></i>
			Showing in &amp; out of stock
			</a>
		<% end if %>
		<% if Request.Querystring("keywords") <> "" then %>
			<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href=""  data-filter="keywords">
			<i class="fa fa-times"></i>
			Keywords: <%= Request.Querystring("keywords") %>
			</a>
		<% end if %>
		<% If request.querystring("new") = "Yes" then %>
			<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href=""  data-filter="new">
			<i class="fa fa-times"></i>
			New items
			</a>
		<% end if %>
		<% If request.querystring("restock") = "Yes" then %>
			<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href=""  data-filter="restock">
			<i class="fa fa-times"></i>
			Restocked items
			</a>
		<% end if %>
		<% If request.querystring("limited") = "yes" then %>
			<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href=""  data-filter="limited"  data-value="<%= request.querystring("limited") %>">
			<i class="fa fa-times"></i>
			Limited items only
			</a>
		<% end if %>
		<% If request.querystring("onetime") = "yes" then %>
			<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href=""  data-filter="onetime"  data-value="<%= request.querystring("onetime") %>">
			<i class="fa fa-times"></i>
			One offs only
			</a>
		<% end if %>
<%
	' Create array for categories
	if request.querystring("jewelry") <> "" then
			build_array = split(request.querystring("jewelry"), ",")
			For i = 0 to Ubound(build_array)
				response.write "<a class=""filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0"" href=""""  data-filter=""jewelry"" data-value=""" + trim(Server.HTMLEncode(build_array(i))) + """><i class=""fa fa-times""></i> " + Server.HTMLEncode(build_array(i)) + "</a>"
			next		
	end if 
	
	' Create array for gauges
	if request.querystring("gauge") <> "" then
			build_array = split(request.querystring("gauge"), ",")
			For i = 0 to Ubound(build_array)
				' Don't take space out of it's a ring size
				if inStr(build_array(i), "Size") > 0 or inStr(build_array(i), "Youth") > 0 or inStr(build_array(i), "Extra") > 0 then
					data_gauge = trim(Server.HTMLEncode(build_array(i)))
				else
					data_gauge =  replace(Server.HTMLEncode(build_array(i)), " " , "")
				end if
				
				response.write "<a class=""filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0"" href="""" data-filter=""gauge"" data-value=""" + data_gauge + """><i class=""fa fa-times""></i> " + Server.HTMLEncode(build_array(i)) + "</a>"
			next		
	end if 

	
	if Request.Querystring("exclude-material") <> "" then
%>
	
		<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href="" data-filter="exclude-material" data-value="<%= request.querystring("exclude-material") %>">
			<i class="fa fa-times"></i>
		Exclude materials ON
		</a>
	<% 
	end if	
	if request.querystring("exclude-material") = "on" then
		var_exclude = " (Excluded)"
	end if
	' Create array for materials
	if request.querystring("material") <> "" then
			build_array = split(request.querystring("material"), ",")
			For i = 0 to Ubound(build_array)
				response.write "<a class=""filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0"" href="""" data-filter=""material"" data-value=""" + trim(Server.HTMLEncode(build_array(i))) + """><i class=""fa fa-times""></i> " + Server.HTMLEncode(build_array(i)) + var_exclude + "</a>"
			next		
	end if
	
	' Create array for brands
	if request.querystring("brand") <> "" then
			build_array = split(request.querystring("brand"), ",")
			For i = 0 to Ubound(build_array)
				response.write "<a class=""filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0"" href="""" data-filter=""brand"" data-value=""" + trim(Server.HTMLEncode(build_array(i))) + """><i class=""fa fa-times""></i> " + Server.HTMLEncode(build_array(i)) + "</a>"
			next		
	end if
	
	' Create array for piercing types
	if request.querystring("piercing") <> "" then
			build_array = split(request.querystring("piercing"), ",")
			For i = 0 to Ubound(build_array)
				response.write "<a class=""filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0"" href="""" data-filter=""piercing"" data-value=""" + trim(Server.HTMLEncode(build_array(i))) + """><i class=""fa fa-times""></i> " + Server.HTMLEncode(build_array(i)) + "</a>"
			next		
	end if
	
	' Create array for flare types
	if request.querystring("flare_type") <> "" then
			build_array = split(request.querystring("flare_type"), ",")
			For i = 0 to Ubound(build_array)
				response.write "<a class=""filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0"" href="""" data-filter=""flare_type"" data-value=""" + Server.HTMLEncode(build_array(i)) + """><i class=""fa fa-times""></i> " + Server.HTMLEncode(build_array(i)) + "</a>"
			next		
	end if

	' Create array for lengths
	if request.querystring("length") <> "" then
			build_array = split(request.querystring("length"), ",")
			For i = 0 to Ubound(build_array)
				response.write "<a class=""filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0"" href="""" data-filter=""length"" data-value=""" + replace(Server.HTMLEncode(build_array(i)), " " , "") + """><i class=""fa fa-times""></i> Length: " + Server.HTMLEncode(build_array(i)) + "</a>"
			next		
	end if 
	
	if Request.Querystring("price") <> "" then
	If Instr(Request.Querystring("price"), ";") Then
		arrPrice = split(request.querystring("price"), ";")	
		if arrPrice(0) > 0 or arrPrice(1) < 100 then
%>
	
		<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href="" data-filter="price" data-value="<%= request.querystring("price") %>">
			<i class="fa fa-times"></i>
		Between <%= FormatCurrency(arrPrice(0), 2) & " - " & FormatCurrency(arrPrice(1), 2) %>
		<%if arrPrice(1)=100 then Response.Write "+"%>
		</a>
	<% 
	end if '=== only display if price filter has been selected
	end if 	'If Request.Querystring("price")
	end if
	
	If request.querystring("discount") <> "" then
	%>
			<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href="" data-filter="discount" data-value="<%= request.querystring("discount") %>">
			<i class="fa fa-times"></i>
			<%= request.querystring("discount") %>% off
			</a>
	<%
	end if ' if sale/discount
	
	' Create array for threading
	if request.querystring("threading") <> "" then
			build_array = split(request.querystring("threading"), ",")
			For i = 0 to Ubound(build_array)
				response.write "<a class=""filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0"" href="""" data-filter=""threading"" data-value=""" + build_array(i) + """><i class=""fa fa-times""></i> " + Server.HTMLEncode(build_array(i)) + "</a>"
			next		
	end if

	If request.querystring("customorders") <> "" and request.querystring("customorders") <> " " then
	%>
			<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href="" data-filter="customorders" data-value="<%= request.querystring("customorders") %>">
			<i class="fa fa-times"></i>
			<% if request.querystring("customorders") = "customorder-not" then %>
				Not showing custom orders
			<% elseif request.querystring("customorders") = "customorder-yes" then %>
				Showing only custom orders
			<% end if %>
			</a>
	<%
	end if
	' Create array for selected colors
	if request.querystring("colors") <> "" then
			if request.querystring("color-filter") = "and" then
			%>
			<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href=""  data-filter="color-filter"  data-value="<%= request.querystring("color-filter") %>">
			<i class="fa fa-times"></i>
			Contains ALL selected colors
			</a>
		<% else %>
		<div class="mt-2 small text-secondary border-bottom border-secondary">Contains ANY selected colors:</div>
	<% end if
	
			build_array = split(request.querystring("colors"), ",")
			For i = 0 to Ubound(build_array)
				response.write "<a class=""filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0"" href="""" data-filter=""colors"" data-value=""" + replace(Server.HTMLEncode(build_array(i)), " " , "") + """><i class=""fa fa-times""></i> " + Server.HTMLEncode(build_array(i)) + "</a>"
			next		
	end if

	If request.querystring("pair") <> "" and request.querystring("pair") <> " "  then
	%>
			<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href="" data-filter="pair" data-value="<%= request.querystring("pair") %>">
			<i class="fa fa-times"></i>
			<% if request.querystring("pair") = "pairs" then %>
				Only pairs
			<% elseif request.querystring("pair") = "singles" then %>
				Only singles
			<% end if %>
			</a>
	<%
	end if	%>
