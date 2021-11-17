<%@LANGUAGE="VBSCRIPT"%>
	<%
	page_title = "Body jewelry | Bodyartforms gauges, septum rings, nose rings & more"
	page_description = "Body jewelry at great prices & HUGE selection of body jewelry! Free jewelry on orders over $30, free o-rings with every order, free basic mail shipping on orders over $50."
	page_keywords = "body jewelry, body piercing jewelry"
	var_extra_head_inc = "homepage"
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
		<link rel="stylesheet" type="text/css" href="/CSS/slick.css"/>
		<!--#include virtual="/bootstrap-template/filters.asp" -->
		<% if request.querystring("status") = "signout" then %>
		<div class="alert alert-success alert-dismissible mb-5">
			<h4>LOGOUT SUCCESSFUL</h4><button type="button" class="close" data-dismiss="alert" aria-label="Close"><span aria-hidden="true">&times;</span></button>

		</div>
		<% end if %>
		<%
SqlString = "SELECT * FROM TBL_Sliders WHERE active= 1 AND GETDATE() BETWEEN date_start AND date_end ORDER BY show_up_order" 
Set rsGetSliders = DataConn.Execute(SqlString)
		
SqlString = "SELECT TOP 20 * FROM FlatProducts WHERE new_page_date >= GETDATE()-45 AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND brandname <> 'Wildcard' AND customorder <> 'yes' ORDER BY new_page_date DESC" 
Set rsGetNewestProducts = DataConn.Execute(SqlString)

SqlString = "SELECT TOP (12) * FROM (SELECT TOP (100) * FROM FlatProducts WHERE active = 1 AND customorder <> 'yes' AND free_item IS NULL AND type = 'None' AND (NOT (material LIKE '%acrylic%')) AND (NOT (tags LIKE '%save%')) ORDER BY qty_sold_last_7_days DESC) as t ORDER BY NEWID()" 
Set rsPopularThisWeek = DataConn.Execute(SqlString)

SqlString = "SELECT TOP(20) Testimonial_Name, Testimonial FROM dbo.TBL_Testimonials WHERE use_on_homepage = 1 ORDER BY newid()" 
Set rsGetTestimonials = DataConn.Execute(SqlString)
%>
<div class="baf-carousel slider-container text-center">
	<div id="HomeSlider">
	<% While NOT rsGetSliders.EOF %>
		<a class="homepage-graphic" href="<%=rsGetSliders("url")%>" id="<%=rsGetSliders("slider_name")%>">
			<picture>
				<source media="(max-width: 550px)" sizes="100vw" srcset="https://sliders.bodyartforms.com/<%=Server.UrlEncode(rsGetSliders("img550x350"))%>">
				<source media="(max-width: 850px)" sizes="100vw" srcset="https://sliders.bodyartforms.com/<%=Server.UrlEncode(rsGetSliders("img850x350"))%>">
				<source media="(max-width: 1024px)" sizes="100vw" srcset="https://sliders.bodyartforms.com/<%=Server.UrlEncode(rsGetSliders("img1024x350"))%>">
				<source media="(max-width: 1600px)" sizes="100vw" srcset="https://sliders.bodyartforms.com/<%=Server.UrlEncode(rsGetSliders("img1600x350"))%>">
				<source media="(min-width: 1920px)" sizes="100vw" srcset="https://sliders.bodyartforms.com/<%=Server.UrlEncode(rsGetSliders("img1920x350"))%>">
					<img class="img-fluid hero-slider" src="<%=Server.UrlEncode(rsGetSliders("img1600x350"))%>" 
				srcset="https://sliders.bodyartforms.com/<%=Server.UrlEncode(rsGetSliders("img1920x350"))%> 1920w,
				https://sliders.bodyartforms.com/<%=Server.UrlEncode(rsGetSliders("img1600x350"))%> 1600w,
				https://sliders.bodyartforms.com/<%=Server.UrlEncode(rsGetSliders("img1024x350"))%> 1024w,
				https://sliders.bodyartforms.com/<%=Server.UrlEncode(rsGetSliders("img850x350"))%> 850w,
				https://sliders.bodyartforms.com/<%=Server.UrlEncode(rsGetSliders("img550x350"))%> 550w"
				sizes="100vw"
				alt="<%=rsGetSliders("slider_name")%>" />
			</picture>
		</a>
		<%rsGetSliders.MoveNext%>	
	<% Wend %>			
	</div>
</div>

			<% 
			If Not rsGetTestimonials.EOF Then %>
			<div class="baf-carousel mb-3" id="testimonials">
			<% 	While NOT rsGetTestimonials.EOF %>
			<div class="slide alert alert-secondary">
				<i class="fa fa-lg fa-double-quote-serif-left pr-2"></i>
				<%=(rsGetTestimonials.Fields.Item("Testimonial").Value)%>
		
				<i class="fa fa-lg fa-double-quote-serif-right pl-2"></i>
				<span class="pl-3">
				</span>
			</div>
			<% rsGetTestimonials.MoveNext()
			Wend
			%>
			</div>
			<% end if 'rsGetTestimonials.EOF 
			 %>

			<a href="/products.asp?new=Yes"><img src="/images/button-shop-new-items.png"></a>


<div class="mt-2 full-width-block">
							<div class="baf-carousel" id="NewSlider">
										<% 
										i = 1
					While NOT rsGetNewestProducts.EOF 
			
					' set variables for pricing
					if rsGetNewestProducts.Fields.Item("min_sale_price").Value <> "" then
						min_price = FormatNumber(rsGetNewestProducts.Fields.Item("min_sale_price").Value,2)
					else
						min_price = ""
					end if
					if rsGetNewestProducts.Fields.Item("max_sale_price").Value <> "" then
						max_price = FormatNumber(rsGetNewestProducts.Fields.Item("max_sale_price").Value,2)
					else
						max_price = ""
					end if
					
					DisplayPrice = ""
					if rsGetNewestProducts.Fields.Item("SaleDiscount").Value > 0 then	
						DisplayPrice = DisplayPrice & rsGetNewestProducts.Fields.Item("SaleDiscount").Value & "% OFF "
					end if
					
					
					if rsGetNewestProducts.Fields.Item("min_sale_price").Value <> "" then
					DisplayPrice = DisplayPrice & "$" & min_price & " "
					end if
					
					if rsGetNewestProducts.Fields.Item("min_sale_price").Value <> rsGetNewestProducts.Fields.Item("max_sale_price").Value then
						DisplayPrice = DisplayPrice & " - $" & max_price
					end if
					%>
												<a class="slide text-dark homepage-graphic" href="productdetails.asp?productid=<%= rsGetNewestProducts.Fields.Item("ProductID").Value %>" id="new-<%= replace(lcase(rsGetNewestProducts.Fields.Item("title").Value)," ", "-") %>-<%=(rsGetNewestProducts.Fields.Item("ProductID").Value)%>">
													<% if i < 9 then %>
													<img class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetNewestProducts.Fields.Item("picture").Value %>" alt="<%=(rsGetNewestProducts.Fields.Item("title").Value)%>" title="<%=(rsGetNewestProducts.Fields.Item("title").Value)%>">
													<% else %><!-- lazy load in images beyond 8-->
													<img class="img-fluid lazyload" src="/images/image-placeholder.png" data-src="https://bafthumbs-400.bodyartforms.com/<%= rsGetNewestProducts.Fields.Item("picture").Value %>" alt="<%=(rsGetNewestProducts.Fields.Item("title").Value)%>" title="<%=(rsGetNewestProducts.Fields.Item("title").Value)%>">
													<% end if 
													i = i +1 %>
												<div class="w-100 text-center px-1">
														<div class="small">
																<%= rsGetNewestProducts.Fields.Item("title").Value %>
														</div>
														<div class="small font-weight-bold  d-block">
																<%= DisplayPrice %>
														</div>
														<div class="small font-weight-bold  d-block">
																<%= rsGetNewestProducts.Fields.Item("min_gauge").Value %>
																<% if rsGetNewestProducts.Fields.Item("min_gauge").Value <> rsGetNewestProducts.Fields.Item("max_gauge").Value then %> 
																- <%= rsGetNewestProducts.Fields.Item("max_gauge").Value %>
																<% end if %>
														</div>
													
													</div> 
											</a>
											<% 
					rsGetNewestProducts.MoveNext()
					Wend
					%>
										</div>
			</div><!-- full-width-block -->

			<div class="container-fluid mt-4">
				<div class="row">
<div class="col-12 d-md-none p-0">		
			<div  class="bg-dark p-3 rounded text-light">
					<div class="p-0 m-0 h5">
							GET NOTIFIED ABOUT SALES!
						</div>
						<div class="small mb-1">Sign up for our newsletter</div>
						<div class="input-group">
								<input type="text" class="form-control bg-lightgrey text-dark border-0 " placeholder="E-mail address" aria-label="E-mail address" type="text" name="homepage_newsletter_email" id="homepage_newsletter_email"  />
								<div class="input-group-append">
								  <button class="btn btn-info bg-info input-group-text px-1 text-white border-0 event-newsletter" type="button" id="homepage-newsletter-signup"><i class="fa fa-paper-plane mr-2"></i> Sign Up</button>
								</div>
							  </div>
							  <div class="mt-1" id="homepage-newsletter-msg"></div>
			</div>
		</div>
		<div class="col-12 d-md-none">
				
		</div>
		</div>
		</div>
<div class="container-fluid mt-4">
	<div class="row justify-content-center">
			<div class="col-12 col-sm-4 col-lg-4 px-1 px-xl-5 mt-2">	
					<a class="homepage-graphic" href="/products.asp?jewelry=nose" id="square-nose-rings"><img class="img-fluid" src="/images/homepage/nose-rings-600px.jpg" alt="banner-nose-rings"></a>
				</div>
		
		<div class="col-12 col-sm-4 col-lg-4 px-1 px-xl-5 mt-2">	
			<a class="homepage-graphic" href="/products.asp?price=20" id="square-under20-graphic"><img class="img-fluid" src="/images/homepage/under-20-bucks.jpg" alt="image"></a>
		</div>
		<div class="col-12 col-sm-4 col-lg-4 px-1 px-xl-5 mt-2">	
				<a class="homepage-graphic" href="/products.asp?jewelry=plugs" id="square-plugs-graphic"><img class="img-fluid" src="/images/homepage/home-plugs-tunnels.jpg" alt="image"></a>
			</div>
	</div><!-- row -->	
</div><!-- main categories fluid container -->



<h1 class="display-5" style="margin-top: 1.5em">Best sellers this week</h1>
<div class="d-flex flex-row flex-wrap">

			<% 
			i = 1
While NOT rsPopularThisWeek.EOF 

' Set variable for date added -------------
var_new_date = ""
if rsPopularThisWeek.Fields.Item("new_page_date").Value <= date()+21 AND rsPopularThisWeek.Fields.Item("new_page_date").Value > date()-70 then
	
	var_new_date = MonthName(Month(rsPopularThisWeek.Fields.Item("new_page_date").Value),1) & " " & Day(rsPopularThisWeek.Fields.Item("new_page_date").Value)
	
end if

' set variables for pricing
if rsPopularThisWeek.Fields.Item("min_sale_price").Value <> "" then
	min_price = FormatNumber(rsPopularThisWeek.Fields.Item("min_sale_price").Value,2)
else
	min_price = ""
end if
if rsPopularThisWeek.Fields.Item("max_sale_price").Value <> "" then
	max_price = FormatNumber(rsPopularThisWeek.Fields.Item("max_sale_price").Value,2)
else
	max_price = ""
end if

DisplayPrice = ""
if rsPopularThisWeek.Fields.Item("SaleDiscount").Value > 0 then	
	DisplayPrice = DisplayPrice & rsPopularThisWeek.Fields.Item("SaleDiscount").Value & "% OFF "
end if


if rsPopularThisWeek.Fields.Item("min_sale_price").Value <> "" then
DisplayPrice = DisplayPrice & "$" & min_price & " "
end if

if rsPopularThisWeek.Fields.Item("min_sale_price").Value <> rsPopularThisWeek.Fields.Item("max_sale_price").Value then
	DisplayPrice = DisplayPrice & " - $" & max_price
end if

' hide anything above 6 from mobile 	
if i > 6 then 
	hide_weekly_best_sellers = "d-none d-md-none d-xl-block"
end if
if i > 8 then 
	hide_weekly_best_sellers = "d-md-none d-none d-xl-block"
end if

%>
<div class="col-6 col-xs-4 col-md-3 col-lg-2 col-xl-2 col-break1600-1 col-break1900-1 my-3 px-1 px-md-2 text-center <%= hide_weekly_best_sellers %>">	
		<div class="container-fluid">
				<div class="row border-bottom border-secondary products">	

						<a class="homepage-graphic" href="productdetails.asp?productid=<%= rsPopularThisWeek.Fields.Item("ProductID").Value %>" id="bestSeller-<%= replace(lcase(rsPopularThisWeek.Fields.Item("title").Value)," ", "-") %>-<%=(rsPopularThisWeek.Fields.Item("ProductID").Value)%>">
									<img class="img-fluid w-100 lazyload" src="/images/image-placeholder.png" data-src="https://bafthumbs-400.bodyartforms.com/<%= rsPopularThisWeek.Fields.Item("picture").Value %>" alt="Product Photo" title="<%=(rsPopularThisWeek.Fields.Item("title").Value)%>">

							</a>

</div><!-- image container end row -->
		<a class="row text-dark" href="productdetails.asp?productid=<%= rsPopularThisWeek.Fields.Item("ProductID").Value %>">
		<div class="w-100 text-left px-1 small">
			<div>
					<%= rsPopularThisWeek.Fields.Item("title").Value %>
			</div>
			<div class="font-weight-bold  d-block">
					<%= DisplayPrice %>
			</div>
			<div class="font-weight-bold  d-block">
					<%= rsPopularThisWeek.Fields.Item("min_gauge").Value %>
					<% if rsPopularThisWeek.Fields.Item("min_gauge").Value <> rsPopularThisWeek.Fields.Item("max_gauge").Value then %> 
					- <%= rsPopularThisWeek.Fields.Item("max_gauge").Value %>
					<% end if %>
			</div>
					<% 
				var_total_reviews = rsPopularThisWeek.Fields.Item("total_reviews").Value
				var_total_photos = rsPopularThisWeek.Fields.Item("total_photos").Value
				'if there are more than 5 ratings then show star ratings
			'	if dont_show = "yes" then
					if rsPopularThisWeek.Fields.Item("avg_rating").Value <> "" then
						var_avg_rating = FormatNumber(rsPopularThisWeek.Fields.Item("avg_rating").Value,1)
						var_avg_percentage = var_avg_rating * 20
					end if ' if there are more than 5 ratings
					
					if rsPopularThisWeek.Fields.Item("avg_rating").Value <> "" then %>
					<div>
							<span class="rating-box">
									<span class="rating" style="width:<%= var_avg_percentage %>%"></span>
								</span>
					<span class="text-dark small ml-3">
					<%= formatnumber(rsPopularThisWeek.Fields.Item("total_reviews").Value,0) %>
				</span>
				</div>
				<% end if %>
			
		</div> 
	</a>
		

</div>	<!-- container-fluid end --> 
</div><!-- flex column -->		
					
			
				<% 
				' hide anything above 6 items on mobile
				i = i + 1
				rsPopularThisWeek.MoveNext()
Wend
%>
</div><!-- popular products flex container -->		
	

			<!--#include virtual="/bootstrap-template/footer.asp" -->
			<script type="text/javascript" src="/js/slick.min.js"></script>
			<script type="text/javascript" src="/js-pages/homepage.min.js?v=111021"></script>