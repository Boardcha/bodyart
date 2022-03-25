<%@ Language=VBScript %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<html>
<head>
<!--#include file="includes/inc_scripts.asp"-->
<title>
	New Products Sorting
</title>

<style>
	#SortableList{
		margin: 0;
		padding: 0;
		text-align: center;
	}
	#SortableList li{
		display: inline-block;
		vertical-align: top;
		width: 200px;
		margin: 10px;
		cursor: url(https://shopify.github.io/draggable/assets/img/cursor-drag.png),auto;
		
	}
	.draggable-source--is-dragging {
		visibility: hidden;
	}
	#SortableList li a:hover{
		cursor: url(https://shopify.github.io/draggable/assets/img/cursor-drag.png),auto;
	}
	.draggable--is-dragging, .draggable--is-dragging * {
		cursor: url(https://shopify.github.io/draggable/assets/img/cursor-drag-clicked.png),auto;
	}
</style>
</head>
<body>
<!--#include file="admin_header.asp"-->
<%
custom_sorting = "yes"
%>
<!--#include virtual="/products/inc_product_search_query.asp"-->

<div class="products">
<h3 class="m-3">New Products Sorting</h3>
	<% if NOT rsGetRecords.EOF then %>

	<ul class="mt-3" id="SortableList">
	
	<% 
	While NOT rsGetRecords.EOF
		counter = 1

		' Set variable for an item that is sold as a pair
		if rsGetRecords.Fields.Item("pair").Value = "yes" then
			DisplayPair = " (pair)"
		Else
			DisplayPair = ""
		End if 

		' set variables for pricing
		if rsGetRecords.Fields.Item("min_sale_price").Value <> "" then
			min_price = FormatNumber(rsGetRecords.Fields.Item("min_sale_price").Value,2)
		else
			min_price = ""
		end if
		if rsGetRecords.Fields.Item("max_sale_price").Value <> "" then
			max_price = FormatNumber(rsGetRecords.Fields.Item("max_sale_price").Value,2)
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
				min_price = FormatNumber(rsGetRecords.Fields.Item("min_price").Value,2)
			else
				min_price = ""
			end if 
			if rsGetRecords.Fields.Item("max_price").Value <> "" then	
				max_price = FormatNumber(rsGetRecords.Fields.Item("max_price").Value,2)
			else
				max_price = ""
			end if 
			
			'--- hide secret sale product from sales page only
			if request.querystring("discount") <> "" then
				hide_product = "d-none"
			end if
		end if

		if rsGetRecords.Fields.Item("min_sale_price").Value <> "" then
			DisplayPrice = DisplayPrice & "$" & min_price & " "
		end if
		
		if rsGetRecords.Fields.Item("min_sale_price").Value <> rsGetRecords.Fields.Item("max_sale_price").Value then
			DisplayPrice = DisplayPrice & " - " & "$" & max_price
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

	<li data-productid="<%= rsGetRecords.Fields.Item("ProductID").Value %>">	
			<div class="container-fluid">
					<div class="row border-bottom border-secondary">	

			<a class="col p-0 text-dark" href="productdetails.asp?ProductID=<%= rsGetRecords.Fields.Item("ProductID").Value %>" data-historyid="nav<%= rsGetRecords.Fields.Item("ProductID").Value %>"><div class="position-relative"><img class="img-fluid w-100" <%= css_inactive %>  src="<%= var_photo_url %>/<%=(rsGetRecords.Fields.Item("picture").Value)%>" title="<%=(rsGetRecords.Fields.Item("title").Value)%>" alt="<%=(rsGetRecords.Fields.Item("title").Value)%>" />
	<%
	if (rsGetRecords.Fields.Item("SaleDiscount").Value > 0 AND rsGetRecords.Fields.Item("secret_sale").Value = 0) OR  (rsGetRecords.Fields.Item("secret_sale").Value = 1 AND session("secret_sale") = "yes") then 
	%>	
	<span class="product-badges badge badge-danger position-absolute rounded-0 p-2 text-left" style="white-space: normal">SALE <%= DisplayPercentageOff %> <s><%= OriginalPrice %></s></span>
	<%
	end if
	%>
				
			<% if DisplayLogo <> "" or DisplayLogoText <> "" then %>
			<div class="brand-info position-absolute w-50 badge badge-light rounded-0">
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
	</li>
	 <%
	 counter = counter + 1
	 rsGetRecords.MoveNext()

	Wend
	%>

	</ul>
		
		<!--#include virtual="/products/inc_product_paging.asp"-->

<% Else ' if no records are found %>
	<h5 class="alert alert-danger mt-3">No results found</h5>
<% End If%>
</div>
<!-- Sortable List -->
<script src="https://cdn.jsdelivr.net/npm/@shopify/draggable@1.0.0-beta.8/lib/draggable.bundle.js"></script>
<script>

	const sortable = new Draggable.Sortable(
		document.querySelector('#SortableList'), {
			draggable: 'li',
	});
	sortable.on('sortable:stop', (e) => {
		updateProductOrder(e)
	});
	
	function updateProductOrder(e){
		var products_sorted = "";
		var ul = document.getElementById("SortableList");
		var items = ul.getElementsByTagName("li");
		for (var i = 0; i < items.length; ++i) {
			if(items[i].style.display != 'none' && items[i].classList.contains('draggable-mirror') == false){
			  products_sorted = products_sorted + ",(" + items[i].getAttribute("data-productid") + ", " + i + ")"
		  }
		}
		console.log(products_sorted.replace(',',''));
		
		$.ajax({
			method: "POST",
			url: "products/ajax_update_new_products_order.asp",
			data: {products: products_sorted.replace(',','')}
		})		
	};	
</script>
</body>
</html>
<%
DataConn.Close()
Set DataConn = Nothing
%>
