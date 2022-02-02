<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<!--#include virtual="/Connections/sql_connection.asp" -->
<%
	
	if (InStr(Keywords, "gift") <> 0 or InStr(Keywords, "gift certificate") <> 0 or InStr(Keywords, "certificate") <> 0) then
		response.redirect "gift-certificate.asp"
	End if	

	
if Request.ServerVariables("QUERY_STRING") <> "" then
' Build targets and break out URL for auto checkbox filling on filters. Also build information from database for meta tags and product page descriptions  ==============================

Function URLDecode(sDec)
	dim objRE
	set objRE = new RegExp
	sDec = Replace(sDec, "+", " ")
	objRE.Pattern = "%([0-9a-fA-F]{2})"
	objRE.Global = True
	URLDecode = objRE.Replace(sDec, GetRef("URLDecodeHex"))
End Function

'// Replacement function for the above
Function URLDecodeHex(sMatch, lhex_digits, lpos, ssource)
	URLDecodeHex = chr("&H" & lhex_digits)
End Function

var_decoded_url = URLDecode(Request.ServerVariables("QUERY_STRING"))

qs_array = Split(var_decoded_url, "&")
seo_array_values = ""

	' If duplicates are found in the URL then remove them from searching and default to generic search results
	if (len(var_decoded_url) - len(replace(var_decoded_url, "jewelry", "")) > 7) and instr(var_decoded_url,"basics") =0 and instr(var_decoded_url,"aftercare") =0 and instr(var_decoded_url,"shield") = 0 then
		no_filter_jewelry = "jewelry"
	else 
		no_filter_jewelry = ""
	end if
	if (len(var_decoded_url) - len(replace(var_decoded_url, "gauge", "")) > 5)  and instr(var_decoded_url,"00g") = 0 then
		no_filter_gauge = "gauge"
	else 
		no_filter_gauge = ""
	end if
	if (len(var_decoded_url) - len(replace(var_decoded_url, "brand", "")) > 5)  then
		no_filter_brand = "brand"
	else 
		no_filter_brand = ""
	end if

	' if we are showing general results due to jewelry and gauge, then remove material as well
	if (no_filter_jewelry <> "" and no_filter_gauge <> "") or len(var_decoded_url) - len(replace(var_decoded_url, "material", "")) > 8 then
		no_filter_material = "material"
	else 
		no_filter_material = ""
	end if

for each x in qs_array
	sub_array = Split(x, "=")

	step_count = 0
	var_qs_name = ""
	var_qs_value = ""
	for each z in sub_array
	seo_value = ""

		if step_count = 0 then
			var_qs_name = "input[name='" + z + "']"
			seo_key = z
			search_key_string = search_key_string & " " &  z
		end if
		
		if step_count = 1 then
		' replace(var_qs_final_build, """","\""")
			var_qs_value = "[value='" + replace(z, """","\""") + "']"

			if seo_key <> "" and seo_key <> no_filter_jewelry and seo_key <> no_filter_gauge and seo_key <> no_filter_material and seo_key <> no_filter_brand and seo_key <> "pagenumber" and seo_key <> "results" and seo_key <> "limited" and seo_key <> "onetime" and seo_key <> "page" and seo_key <> "gauge2" and seo_key <> "gauge3" and seo_key <> "jewelry2" and seo_key <> "jewelry3" and seo_key <> "colors" and seo_key <> "price" and seo_key <> "length" and seo_key <> "preorders" and seo_key <> "pair" and seo_key <> "RecordDisplay" and seo_key <> "date" and seo_key <> "Filter" and seo_key <> "Retainer" and seo_key <> "dwzPage" and seo_key <> "void" and seo_key <> "orderby" and seo_key <> "restock" and seo_key <> "more" then
				seo_value = z
				if seo_value = "1""" or seo_value = "2""" or seo_value = "3""" then
					seo_value = replace(seo_value, """", "quot")
				end if
				' code out finger rings	
				if seo_value = "rings" then
					seo_value = replace(seo_value, "rings", "finger-ring")
				end if 
				seo_array_values = seo_array_values & "," & replace(replace(seo_value, " ", "+"), "/", "2F")
				'response.write seo_key & ":" & seo_value & ", "
			end if
			if seo_key = "new" and z = "Yes" then
				seo_array_values = seo_array_values & " new "
				'response.write "<br/>" & seo_array_values
			end if
			if seo_key = "discount" and z = "all" then
				seo_array_values = seo_array_values & " sale "
			end if
			
		end if
		var_qs_targets = "#form-filters " & var_qs_name & var_qs_value & ", "
	'	response.write "<br/>" & var_qs_targets
		step_count = step_count + 1
	next
	
	var_qs_final_build = var_qs_final_build & var_qs_targets	
next	
' Have to do -2 because there is a space after the comma
var_qs_final_build = left(var_qs_final_build,len(var_qs_final_build)-2)


if seo_array_values <> "" then
	' remove quotes, replace commas with or for full text search
	var_seo_parameter = replace(replace(Right(seo_array_values, Len(seo_array_values)-1), ",", " "), """","")

	' assign as gauge only if that's the only thing in the querystring
	' for link in gauge drop down navigation only
	if InStr(var_decoded_url, "gauge=") > 0 AND InStr(var_decoded_url, "=plugs") = 0 then
		var_seo_parameter = var_seo_parameter & " onlygauge "
	end if

	if trim(var_seo_parameter) = "" then
		var_seo_parameter = "id=0"
	end if
else
	var_seo_parameter = "id=0"
end if

if var_seo_parameter = "" then ' for when bots or people have querystrings where the key = nothing (ie jewelry=)
	var_seo_parameter = "id=0"
end if

	'response.write "<br/>FINAL BUILD: " & var_qs_final_build
	'response.write "<br>Values: -" & var_seo_parameter & "-"
	'response.write "<br/>Search string: " & search_key_string
	'response.write "<br/>" & "SELECT TOP(1) * FROM tbl_sitemap_searches AS FT_TBL INNER JOIN FREETEXTTABLE( tbl_sitemap_searches, (extra_keywords), " & var_seo_parameter & ", 1) AS KEY_TBL ON FT_TBL.id = KEY_TBL.[KEY] ORDER BY rank DESC"


	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TOP(1) * FROM tbl_sitemap_searches AS FT_TBL INNER JOIN FREETEXTTABLE(tbl_sitemap_searches, (extra_keywords), ?, 1) AS KEY_TBL ON FT_TBL.id = KEY_TBL.[KEY] ORDER BY rank DESC"
	objCmd.Parameters.Append(objCmd.CreateParameter("url",200,1,500, var_seo_parameter))
	set rsSiteMap = objCmd.Execute()

	if NOT rsSiteMap.eof then
		title =  rsSiteMap.Fields.Item("meta_title").Value
		title_onpage =  rsSiteMap.Fields.Item("meta_title_onpage").Value
		description = rsSiteMap.Fields.Item("meta_description").Value
	else
		title = "Search Results"	
		title_onpage = "Search Results"
		description = "Search Results"
	end if'


end if ' Request.ServerVariables("QUERY_STRING") <> ""

page_title = title
page_description = description
page_keywords = ""
var_meta_products_aggregate = "yes"
%>
<!--#include virtual="/products/inc_product_search_query.asp"-->
<% if NOT rsGetRecords.EOF then

If IsNumeric(Request.Querystring("pagenumber")) Then
	if Request.Querystring("pagenumber") = "" then
		CurrentPage = 1
	else
		temp_pagenumber = cint(Request.Querystring("pagenumber"))
	end if

	If temp_pagenumber = 0 or temp_pagenumber > cint(TotalPages) Then
		CurrentPage = 1
	Else
		CurrentPage = Request.Querystring("pagenumber")
	End if
else 
	CurrentPage = 1
end if ' check if pagenumber is numeric

' Product paging counts
' clean original querystring URL
var_qs_url = replace(Request.ServerVariables("QUERY_STRING"), "&pagenumber=0", "")
' clean querystring if coming in from site with pagenumber that is too high
if temp_pagenumber > cint(TotalPages) Then
	var_qs_url = replace(var_qs_url, "&pagenumber=" & Request.Querystring("pagenumber"), "")
end if


' Retrieve page numbers BEFORE current page
	' Get current page - 2
	var_lowest_page = CurrentPage - 2
	' If lowest page is below 0, set the page number to 1
	if var_lowest_page <= 0 then
		var_lowest_page = 1
	end if

' Retrieve page numbers AFTER current page
	' Get current page + 2
	var_highest_page = CurrentPage + 2
	' If highest page is above the total pages, then set it to the total pages
	if var_highest_page >= TotalPages then
		var_highest_page = TotalPages
	end if

	var_begin_results_count = (CurrentPage - 1) * session("resultsperpage") + 1
	var_end_results_count = CurrentPage * session("resultsperpage")

end if ' NOT rsGetRecords.EOF
	%>
<!-- store session variable for continue shopping button only when on product pages -->
<!--#include virtual="/cart/inc-continue-shopping-button.asp" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<%
' ===========================================================================
' set addon cookie if coming in from a link
' ===========================================================================

if request.querystring("addon") = "yes" then
	response.cookies("OrderAddonsActive") = request.querystring("id")
end if
%>
<!--#include virtual="/bootstrap-template/filters.asp" -->

<div class="products">
<div class="display-5 mb-3" style="font-size:1.6em">
	<%= title_onpage %>
</div>

<div class="top-filters">

<% if CustID_Cookie <> 0 and (request.querystring("pagenumber") = "" or request.querystring("pagenumber") = "1") then %>
<div class="my-3">
	<button class="btn-save-search btn btn-sm btn-outline-info text-center"><i class="fa fa-heart"></i> Save this search to my account</button>
	<input type="hidden" name="save-search-string" id="save-search-string" value="<%= Request.ServerVariables("QUERY_STRING") %>">
</div>
<% end if %>

<% If NOT rsGetRecords.EOF then %>
<form name="frm-sort" id="frm-sort" method="post" action="products.asp?<%= replace(Request.ServerVariables("QUERY_STRING"), request.querystring("pagenumber"), "1") %>">
	<%
	If Session("filter_orderby") = "new_page_date desc" then
	ProductSortText = "Newest first"
	Elseif  Session("filter_orderby") = "min_sale_price asc" then
	ProductSortText = "Lowest price first"
	Elseif  Session("filter_orderby") = "min_sale_price desc" then
	ProductSortText = "Highest price first"
	Elseif  Session("filter_orderby") = "avg_rating desc" then
	ProductSortText = "Highest rated"
	Elseif  Session("filter_orderby") = "total_reviews desc" then
	ProductSortText = "Most reviews"
	Elseif  Session("filter_orderby") = "total_photos desc" then
	ProductSortText = "Most customer photos"
	Elseif  Session("filter_orderby") = "title ASC" then
	ProductSortText = "Product title (A-Z)"
	Elseif  Session("filter_orderby") = "material ASC" then
	ProductSortText = "Material"
	Elseif  Session("filter_orderby") = "qty_sold_last_7_days desc" then
	ProductSortText = "Top sellers this week"
	Else
	ProductSortText = ""
	End if
	%>
	<% if request("feature") <> "top_seller" then %>
	<div class="dropdown d-inline-block my-2 my-sm-0  mr-sm-2">
		<button class="btn btn-sm btn-outline-secondary text-left dropdown-toggle" type="button" id="dropdownSort" data-flip="false" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
		  SORT BY: <%= ProductSortText %>
		</button>
		<div class="dropdown-menu modal-scroll-long" style="min-width:14rem" aria-labelledby="dropdownSort">
			<div class="dropdown-item btn-group-vertical btn-group-toggle m-0 p-0 " data-toggle="buttons">

		<label class="btn btn-light d-block text-left">
			<input type="radio" name="sortby" value="new_page_date desc">Newest first (default)
		</label> 
		<label class="btn btn-light d-block text-left">
				<input type="radio" name="sortby" value="min_sale_price asc">Lowest price first
			</label> 
			<label class="btn btn-light d-block text-left">
					<input type="radio" name="sortby" value="min_sale_price desc">Highest price first
				</label> 
				<label class="btn btn-light d-block text-left">
						<input type="radio" name="sortby" value="avg_rating desc">Highest rated
					</label> 
	<% if request.querystring("limited") = "" and request.querystring("onetime") = "" then %>
					<label class="btn btn-light d-block text-left">
							<input type="radio" name="sortby" value="qty_sold_last_7_days desc">Top sellers this week
						</label> 
	<% end if %>
					<label class="btn btn-light d-block text-left">
							<input type="radio" name="sortby" value="total_reviews desc">Most reviews
						</label> 
						<label class="btn btn-light d-block text-left">
								<input type="radio" name="sortby" value="total_photos desc">Most customer photos
							</label> 
							<label class="btn btn-light d-block text-left">
									<input type="radio" name="sortby" value="material ASC">Material
								</label> 

</div><!-- button group -->
</div><!-- drop down menu -->
</div><!-- drop down -->
<% end if 'display if not top sellers category %>
	<div class="dropdown d-inline-block my-2 my-sm-0 mr-sm-2">
			<button class="btn btn-sm btn-outline-secondary text-left dropdown-toggle" type="button" id="dropdownResults" data-flip="false" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
					Viewing <% if TotalRecords = Session("resultsperpage") then %>all (<%= TotalRecords %>)<%else%><%= Session("resultsperpage") %> per page<% end if %>
			</button>
			<div class="dropdown-menu modal-scroll-long" aria-labelledby="dropdownResults">
				<div class="dropdown-item btn-group-vertical btn-group-toggle m-0 p-0 " data-toggle="buttons">
	
			<label class="btn btn-light d-block text-left">
				<input type="radio" name="resultsperpage" value="25">25
			</label> 
			<label class="btn btn-light d-block text-left">
					<input type="radio" name="resultsperpage" value="50">50
				</label> 
				<label class="btn btn-light d-block text-left">
						<input type="radio" name="resultsperpage" value="75">75
					</label> 
					<label class="btn btn-light d-block text-left">
							<input type="radio" name="resultsperpage" value="100">100
						</label> 
						<label class="btn btn-light d-block text-left">
								<input type="radio" name="resultsperpage" value="150">150
							</label> 
							<label class="btn btn-light d-block text-left">
									<input type="radio" name="resultsperpage" value="200">200
								</label> 
								<% if TotalRecords < 500 then %>
								<label class="btn btn-light d-block text-left">
										<input type="radio" name="resultsperpage" value="<%= TotalRecords %>">View All
									</label> 
									<% end if
		%>
	
	</div><!-- button group -->
	</div><!-- drop down menu -->
	</div><!-- drop down -->

		<% if request.cookies("product-display") = "" then 
		varMobGridCols = "col-6"
		%>
		<span class="btn btn-sm btn-outline-secondary ml-2 d-xs-none selector-product-display" data-display="list"><i class="fa fa-list fa-lg"></i></span>
		<% else
			varMobGridCols = "col-12" 
		%>
		<span class="btn btn-sm btn-outline-secondary ml-2 d-xs-none selector-product-display" data-display="grid"><i class="fa fa-grid fa-lg"></i></span><% end if %>

	<a class="btn btn-secondary btn-sm text-white d-lg-none my-1 d-block d-sm-inline-block d-lg-none ml-md-2" id="refine-results" data-toggle="collapse" data-target="#filters" aria-controls="filters" aria-expanded="false" aria-label="Toggle filters"><i class="fa fa-filter mr-2"></i> Refine results</a>
	</form>
	<div class="d-lg-none mt-1 mt-lg-0">
	 <!--#include virtual="/products/inc-selected-filters.asp"-->
	</div>
	<% end if 'NOT rsGetRecords.EOF %>
</div><!-- top-filters -->	


<!--#include virtual="/products/inc-landing-links.asp" -->
<% if NOT rsGetRecords.EOF then
	%>
	<% if TotalPages > 1 then ' Only show pagination if there are more than 1 page %>
	<div class="small mt-3 mb-1 text-center">
		<%= TotalRecords %> Items<span class="mob-hide"> found </span> (Viewing <%= var_begin_results_count %> - <% if var_end_results_count > TotalRecords then %><%= TotalRecords %><% else %><%= var_end_results_count %><% end if %>)
	</div>
		<% end if %>
<!--#include virtual="/products/inc_product_paging.asp"-->



<div class="d-flex flex-row flex-wrap">
<% 
if NOT rsGetRecords.EOF then
lazy_count = 1
	rsGetRecords.AbsolutePage = CurrentPage '======== PAGING
	For intRecord = 1 To rsGetRecords.PageSize

	' Set variable for an item that is sold as a pair
	if rsGetRecords.Fields.Item("pair").Value = "yes" then
		DisplayPair = " (pair)"
	Else
		DisplayPair = ""
	End if 
	
	' set variables for pricing
	if rsGetRecords.Fields.Item("min_sale_price").Value <> "" then
		min_price = FormatNumber(rsGetRecords.Fields.Item("min_sale_price").Value * exchange_rate,2)
	else
		min_price = ""
	end if
	if rsGetRecords.Fields.Item("max_sale_price").Value <> "" then
		max_price = FormatNumber(rsGetRecords.Fields.Item("max_sale_price").Value * exchange_rate,2)
	else
		max_price = ""
	end if

DisplayPrice = ""
OriginalMax = ""
OriginalPrice = ""
hide_product = ""
if (rsGetRecords.Fields.Item("SaleDiscount").Value > 0 AND rsGetRecords.Fields.Item("secret_sale").Value = 0) OR  (rsGetRecords.Fields.Item("secret_sale").Value = 1 AND session("secret_sale") = "yes") then 
		DisplayPercentageOff = rsGetRecords.Fields.Item("SaleDiscount").Value & "% OFF"

			if rsGetRecords.Fields.Item("max_price").Value <> rsGetRecords.Fields.Item("min_price").Value then
				OriginalMax = " to $" & rsGetRecords.Fields.Item("max_price").Value
			end if

		

if rsGetRecords.Fields.Item("min_price").Value <> "" then 
		OriginalPrice = "$" & FormatNumber(rsGetRecords.Fields.Item("min_price").Value,2) & OriginalMax
	end if

	else ' show regular prices
		if rsGetRecords.Fields.Item("min_price").Value <> "" then	
			min_price = FormatNumber(rsGetRecords.Fields.Item("min_price").Value * exchange_rate,2)
		else
			min_price = ""
		end if 
		if rsGetRecords.Fields.Item("max_price").Value <> "" then	
			max_price = FormatNumber(rsGetRecords.Fields.Item("max_price").Value * exchange_rate,2)
		else
			max_price = ""
		end if 
		
		'--- hide secret sale product from sales page only
		if request.querystring("discount") <> "" then
			hide_product = "d-none"
		end if
	end if

if rsGetRecords.Fields.Item("min_sale_price").Value <> "" then
	DisplayPrice = DisplayPrice & exchange_symbol & min_price & " "
end if
	
	if rsGetRecords.Fields.Item("min_sale_price").Value <> rsGetRecords.Fields.Item("max_sale_price").Value then
		DisplayPrice = DisplayPrice & " - " & exchange_symbol & max_price
	end if
	
	DisplayLogoText = ""
	DisplayLogo = ""
	
		if rsGetRecords.Fields.Item("ProductLogo").Value <> "" then
			DisplayLogoText =  ""
			if rsGetRecords.Fields.Item("ShowTextLogo").Value = "Y" then
				img_alt_brand =  rsGetRecords.Fields.Item("brandname").Value
			end if

			DisplayLogo = "<img class=""img-fluid"" src=""images/" & rsGetRecords.Fields.Item("ProductLogo").Value &""" alt=""brand-logo-" & img_alt_brand & """ />"
			
			
		else
		
			if rsGetRecords.Fields.Item("ShowTextLogo").Value = "Y" then
			
				DisplayLogoText =  rsGetRecords.Fields.Item("brandname").Value
				
			end if
		end if

	
	' Set variable for custom orders ----------------
	var_title = ""
	if InStr( 1, (rsGetRecords.Fields.Item("title").Value), "CUSTOM ORDER", vbTextCompare ) then
		
		var_title = Replace(rsGetRecords.Fields.Item("title").Value, "CUSTOM ORDER", "<span class='badge badge-secondary' style='font-size:.9em'>CUSTOM ORDER </span>")
		
	else
	
		var_title = (rsGetRecords.Fields.Item("title").Value)
		
	end if
	

	'Set variable for more colors ---------------
	var_colors = ""
	color_black = ""
	color_blue = ""
	color_red = ""
	color_purple = ""
	color_green = ""
	color_pink = ""
	color_brown = ""
	color_orange = ""
	color_yellow = ""
	color_gold = ""
	color_rosegold = ""
	color_silver = ""
	color_teal = ""
	color_turquoise = ""
	if rsGetRecords.Fields.Item("colors").Value > 1 then
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"black") > 0 then
			color_black = "<span class=""swatch-circle"" style=""background-color:#000""></span>"
		end if
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"blue") > 0 then
			color_blue = "<span class=""swatch-circle"" style=""background-color:#0431B4""></span>"
		end if
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"red") > 0 then
			color_red = "<span class=""swatch-circle"" style=""background-color:#B40404""></span>"
		end if
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"purple") > 0 then
			color_purple = "<span class=""swatch-circle"" style=""background-color:#5F04B4""></span>"
		end if
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"green") > 0 then
			color_green = "<span class=""swatch-circle"" style=""background-color:#088A08""></span>"
		end if
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"pink") > 0 then
			color_pink = "<span class=""swatch-circle"" style=""background-color:#F781D8""></span>"
		end if
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"brown") > 0 then
			color_brown = "<span class=""swatch-circle"" style=""background-color:#8A4B08""></span>"
		end if
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"orange") > 0 then
			color_orange = "<span class=""swatch-circle"" style=""background-color:#FFBF00""></span>"
		end if
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"yellow") > 0 then
			color_yellow = "<span class=""swatch-circle"" style=""background-color:#F4FA58""></span>"
		end if
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"teal") > 0 then
			color_teal = "<span class=""swatch-circle"" style=""background-color:#008080""></span>"
		end if
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"turquoise") > 0 then
			color_turquoise = "<span class=""swatch-circle"" style=""background-color:#40E0D0""></span>"
		end if
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"gold") > 0 then
			color_gold = "<span class=""swatch-circle"" style=""background: linear-gradient(135deg, rgba(254,252,234,1) 0%,rgba(249,208,57,1) 100%)""></span>"
		end if
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"rose-gold") > 0 then
			color_rosegold = "<span class=""swatch-circle"" style=""background: linear-gradient(135deg, rgba(252,217,217,1) 0%,rgba(252,184,184,1) 45%,rgba(255,137,137,1) 100%)""></span>"
		end if
		if instr(rsGetRecords.Fields.Item("color_tags").Value,"silver") > 0 or instr(rsGetRecords.Fields.Item("color_tags").Value,"gray") > 0 then
			color_silver = "<span class=""swatch-circle"" style=""background: linear-gradient(135deg, rgba(238,238,238,1) 0%,rgba(204,204,204,1) 100%);""></span>"
		end if
		

		' Does not show color charts if there is not multiple colored items to choose from
		var_colors = "<div class=""row""><div class=""col h5 p-0 m-0"" style=""overflow-wrap: break-word"">" & color_black & color_blue & color_red & color_purple & color_green & color_pink & color_brown & color_orange & color_yellow & color_gold & color_rosegold & color_silver & color_teal & color_turquoise & "</div></div>"
	end if	'	 colors > 1
	
	
		var_photo_url = "https://bafthumbs-400.bodyartforms.com"
		css_inactive = ""
	' if admin user style inactive products & show photos
	if request.cookies("adminuser") = "yes" and rsGetRecords.Fields.Item("active").Value = 0 then
		var_photo_url = "https://bodyartforms-products.bodyartforms.com"
		css_inactive = " style=""border: 10px solid #FA5858"""
	else
		var_photo_url = "https://bafthumbs-400.bodyartforms.com"
		css_inactive = ""
	end if
	%>

<div class="<%= varMobGridCols %> col-xs-4 col-md-3 col-xl-3 col-break1600-2 my-3 px-1 px-md-2 text-center <%= hide_product %>">	
		<div class="container-fluid">
				<div class="row border-bottom border-secondary">	

		<a class="col p-0 text-dark" href="productdetails.asp?ProductID=<%= rsGetRecords.Fields.Item("ProductID").Value %>" data-historyid="nav<%= rsGetRecords.Fields.Item("ProductID").Value %>"><div class="position-relative"><img class="img-fluid w-100 <% if lazy_count > 20 then %> lazyload <% end if %>" <%= css_inactive %>  <% if lazy_count > 20 then %> src="/images/image-placeholder.png" data-src="<%= var_photo_url %>/<%=(rsGetRecords.Fields.Item("picture").Value)%>" <% else %> src="<%= var_photo_url %>/<%=(rsGetRecords.Fields.Item("picture").Value)%>" <% end if %> title="<%=(rsGetRecords.Fields.Item("title").Value)%>" alt="<%=(rsGetRecords.Fields.Item("title").Value)%>" />
<%
if (rsGetRecords.Fields.Item("SaleDiscount").Value > 0 AND rsGetRecords.Fields.Item("secret_sale").Value = 0) OR  (rsGetRecords.Fields.Item("secret_sale").Value = 1 AND session("secret_sale") = "yes") then 
%>	
<span class="product-badges badge badge-danger position-absolute rounded-0 p-2 text-left" style="white-space: normal">SALE <%= DisplayPercentageOff %> <s><%= OriginalPrice %></s></span>
<%
end if
%>
			
		<% if DisplayLogo <> "" or DisplayLogoText <> "" then %>
		<div class="brand-info position-absolute w-50 badge badge-light rounded-0">
			<%= DisplayLogo %><%= DisplayLogoText %>	
		</div>
		<% end if %>
		</div><!-- position-relative -->
		
	</a> 
</div><!-- image container end row -->
<a class="text-dark" href="productdetails.asp?ProductID=<%= rsGetRecords.Fields.Item("ProductID").Value %>" data-historyid="nav<%= rsGetRecords.Fields.Item("ProductID").Value %>">
<div class="row">
		<div class="small text-center w-100  px-1"><%= var_title %></div> 
		</div>


	<% if Request.Querystring("restock") <> "restock" then %>
	
			
				<div class="row">
					<div class="text-center w-100 font-weight-bold"><%= DisplayPrice %><%= DisplayPair %></div>
				</div>
				<% if rsGetRecords.Fields.Item("min_gauge").Value <> "" and rsGetRecords.Fields.Item("min_gauge").Value <> "n/a" and rsGetRecords.Fields.Item("min_gauge").Value <> "nbsp;" and rsGetRecords.Fields.Item("min_gauge").Value <> " " then %>
				<div class="row">
					<div class="text-center w-100 font-weight-bold ">
				<%= rsGetRecords.Fields.Item("min_gauge").Value %>
						<% if rsGetRecords.Fields.Item("min_gauge").Value <> rsGetRecords.Fields.Item("max_gauge").Value then %> 
						- <%= rsGetRecords.Fields.Item("max_gauge").Value %>
						<% end if %>
					</div>
				</div><!-- end row -->
				<% end if %>
			
		
	
	
	<!-- more colors container row -->
			<%= var_colors %>
			
	
		
		<div class="row">
			<div class="col-12 mb-1 text-center px-0">
		<% 
		var_total_reviews = rsGetRecords.Fields.Item("total_reviews").Value
		var_total_photos = rsGetRecords.Fields.Item("total_photos").Value
		'if there are more than 5 ratings then show star ratings
	'	if dont_show = "yes" then
			if rsGetRecords.Fields.Item("avg_rating").Value <> "" then
				var_avg_rating = FormatNumber(rsGetRecords.Fields.Item("avg_rating").Value,1)
				var_avg_percentage = var_avg_rating * 20
			end if ' if there are more than 5 ratings
			
			if rsGetRecords.Fields.Item("avg_rating").Value <> "" then %>
			<span class="rating-box">
					<span class="rating" style="width:<%= var_avg_percentage %>%"></span>
				</span>
				<% end if %>

				<% if var_total_photos > 0 then %>
				<a class="ml-3 small text-dark" href="productdetails.asp?ProductID=<%= rsGetRecords.Fields.Item("ProductID").Value %>#photos"><%= var_total_photos %> photos</a>
				<% end if %>
			</div>
			</div><!-- reviews and photo counts -->
				
	<% else ' if jewelry is restocked %>
	
			<div class="font-weight-bold">
			<%= rsGetRecords.Fields.Item("gauge").Value %>&nbsp;<%= rsGetRecords.Fields.Item("length").Value %>&nbsp;<%= rsGetRecords.Fields.Item("ProductDetail1").Value %>
		</div>
		<div class="font-weight-bold">
			<%= FormatCurrency(rsGetRecords.Fields.Item("price").Value,2) %>
		</div>
		<div class="small">
			Re-stocked: <%= rsGetRecords.Fields.Item("Daterestocked").Value %></div>
		

	<% end if ' restock show %>  
		</a>
</div>	<!-- container-fluid end --> 
</div><!-- flex column -->
 <%
 lazy_count = lazy_count + 1
  rsGetRecords.MoveNext()
If rsGetRecords.EOF Then Exit For  ' ====== PAGING
Next ' ====== PAGING

end if ' if recordset not empty
%>

</div><!-- d-flex -->
	
	<!--#include virtual="/products/inc_product_paging.asp"-->

<% else ' if no records are found %>
		<h5 class="alert alert-danger mt-3">No results found</h5>
<% end if 'NOT rsGetRecords.EOF %>	

	

</div>


	<button class="products-top rounded-circle text-center position-fixed px-2 py-1 alert alert-secondary pointer" type="button"><i class="fa fa-chevron-up"></i></button>
</div><!-- end main content-box -->
<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript">	
	// Check all selected box from the querystring
	
	$("<%= var_qs_final_build %>").prop('checked', true);
	
//	console.log("FORM VALUES: " + $("#form-filters :input[value!='']").serialize());
	</script>
<script type="text/javascript" src="/js-pages/currency-exchange.min.js?v=050619"></script>
<% if (session("exchange-rate") = "" OR session("exchange-currency") <> request.cookies("currency")) AND request.cookies("currency") <> "" AND request.cookies("currency") <> "USD" then %>
<script>
		// Get currency conversions on page load
		updateCurrency();
</script>
<% end if %>
<script type="text/javascript" src="/js-pages/products.min.js?v=040221"></script>