<%@LANGUAGE="VBSCRIPT"%>
	<%
	page_title = "Threadless body jewelry"
	page_description = "Quality threadless jewelry. Make your life easier by avoiding tiny threaded jewelry slipping out of your fingers."
	page_keywords = "threadless jewelry"
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<link rel="stylesheet" type="text/css" href="/CSS/slick.css"/>
<!--#include virtual="/bootstrap-template/filters.asp" -->


<h4>How threadless body jewelry works</h4>
What is threadless jewelry? Give a description and why it's great.

<div class="container-fluid">
	<div class="row">
		<div class="col-xl-6 col-12">
			<video class="mw-100" width="560" height="315" preload="metadata" controls muted>
				<source src="https://videos.bodyartforms.com/video-threadless-ends-how-to.mp4#t=0.5" type="video/mp4">
			  Your browser does not support playing embedded videos
			  </video>
		</div>
	</div>
	<div class="row">
		<div class="col-lg-3 col-6 rounded" style="background-color: #E6E6E6">
			<img class="img-fluid" src="/images/threadless-end-01.png">
			<div>text</div>
		</div>	
		<div class="col-lg-3 col-6 rounded mx-1" style="background-color: #E6E6E6">
			<img class="img-fluid" src="/images/threadless-end-02.png">
			<div>text</div>
		</div>	
		<div class="col-lg-3 col-6 rounded mx-1" style="background-color: #E6E6E6">
			<img class="img-fluid" src="/images/threadless-end-03.png">
			<div>text</div>
		</div>	
		<div class="col-lg-3 col-6 rounded" style="background-color: #E6E6E6">
			<img class="img-fluid" src="/images/threadless-end-04.png">
			<div>text</div>
		</div>	
	</div><!-- row -->
</div><!-- container-->

<h4>Threadless Best Sellers</h4>
  <div class="mt-2 full-width-block">
	<div class="baf-carousel" id="BestSellers">
				<% 

				SqlString = "SELECT TOP 20 * FROM FlatProducts WHERE tags LIKE '%threadless%' AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC" 
				Set rsGetThreadless = DataConn.Execute(SqlString)


				i = 1
					While NOT rsGetThreadless.EOF 

					' set variables for pricing
					if rsGetThreadless.Fields.Item("min_sale_price").Value <> "" then
					min_price = FormatNumber(rsGetThreadless.Fields.Item("min_sale_price").Value,2)
					else
					min_price = ""
					end if
					if rsGetThreadless.Fields.Item("max_sale_price").Value <> "" then
					max_price = FormatNumber(rsGetThreadless.Fields.Item("max_sale_price").Value,2)
					else
					max_price = ""
					end if

					DisplayPrice = ""
					if rsGetThreadless.Fields.Item("SaleDiscount").Value > 0 then	
					DisplayPrice = DisplayPrice & rsGetThreadless.Fields.Item("SaleDiscount").Value & "% OFF "
					end if


					if rsGetThreadless.Fields.Item("min_sale_price").Value <> "" then
					DisplayPrice = DisplayPrice & "$" & min_price & " "
					end if

					if rsGetThreadless.Fields.Item("min_sale_price").Value <> rsGetThreadless.Fields.Item("max_sale_price").Value then
					DisplayPrice = DisplayPrice & " - $" & max_price
					end if
					%>
						<a class="slide text-dark homepage-graphic" href="productdetails.asp?productid=<%= rsGetThreadless.Fields.Item("ProductID").Value %>" id="new-<%= replace(lcase(rsGetThreadless.Fields.Item("title").Value)," ", "-") %>-<%=(rsGetThreadless.Fields.Item("ProductID").Value)%>">
							<% if i < 9 then %>
							<img class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetThreadless.Fields.Item("picture").Value %>" alt="<%=(rsGetThreadless.Fields.Item("title").Value)%>" title="<%=(rsGetThreadless.Fields.Item("title").Value)%>">
							<% else %><!-- lazy load in images beyond 8-->
							<img class="img-fluid lazyload" src="/images/image-placeholder.png" data-src="https://bafthumbs-400.bodyartforms.com/<%= rsGetThreadless.Fields.Item("picture").Value %>" alt="<%=(rsGetThreadless.Fields.Item("title").Value)%>" title="<%=(rsGetThreadless.Fields.Item("title").Value)%>">
							<% end if 
							i = i +1 %>
						<div class="w-100 text-center px-1">
								<div class="small">
										<%= rsGetThreadless.Fields.Item("title").Value %>
								</div>
								<div class="small font-weight-bold  d-block">
										<%= DisplayPrice %>
								</div>
								<div class="small font-weight-bold  d-block">
										<%= rsGetThreadless.Fields.Item("min_gauge").Value %>
										<% if rsGetThreadless.Fields.Item("min_gauge").Value <> rsGetThreadless.Fields.Item("max_gauge").Value then %> 
										- <%= rsGetThreadless.Fields.Item("max_gauge").Value %>
										<% end if %>
								</div>
							
							</div> 
					</a>
					<% 
rsGetThreadless.MoveNext()
Wend
%>
				</div>
</div><!-- full-width-block -->


<div class="container-fluid">

	<div class="col-6">
		<a href="/products.asp?threading=Threadless">Link to all threadless jewelry</a>
	</div>
	<div class="row text-center py-3">
		<div class="col-6">
			<a href="/products.asp?jewelry=balls&threading=Threadless">Threadless ends</a>
		</div>
		<div class="col-6">
			<a href="/products.asp?jewelry=labret&threading=Threadless">Threadless labrets</a>
		</div>
	
		<div class="col-6">
			<a href="/products.asp?jewelry=nose-ring&price=0%3B100&threading=Threadless">Threadless nose jewelry</a>
		</div>
		<div class="col-6">
			<a href="/products.asp?jewelry=chains-necklace">Threadless barbells</a>
		</div>
		<div class="col-6">
			<a href="/products.asp?jewelry=curved&threading=Threadless">Threadless curved barbells</a>
		</div>
		<div class="col-6">
			<a href="/products.asp?jewelry=barbell&threading=Threadless">Threadless straight barbells</a>
		</div>
	</div>

</div><!-- container -->


<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript" src="/js/slick.min.js"></script>
<script type="text/javascript">
	$('#BestSellers').slick({
	slidesToShow: 3,
  slidesToScroll: 3,
  prevArrow: '<div class="slider-arrow-prev" style="height:60%"><i class="fa fa-chevron-circle-left fa-2x text-dark pointer"></i></div>',
  nextArrow: '<div class="slider-arrow-next" style="height:60%"><i class="fa fa-chevron-circle-right fa-2x text-dark pointer"></i></div>',
  responsive: [

    {
      breakpoint: 4000,
      settings: {
        slidesToShow: 10,
        slidesToScroll: 10
      }
    },
    {
      breakpoint: 1920,
      settings: {
        slidesToShow: 8,
        slidesToScroll: 8
      }
    },
    {
      breakpoint: 1600,
      settings: {
        slidesToShow: 7,
        slidesToScroll: 7
      }
    },
    {
      breakpoint: 1024,
      settings: {
        slidesToShow: 5,
        slidesToScroll: 5
      }
    },
    {
      breakpoint: 600,
      settings: {
        slidesToShow: 4,
        slidesToScroll: 4
      }
    },
    {
      breakpoint: 480,
      settings: {
        slidesToShow: 3,
        slidesToScroll: 3
      }
    }
    // You can unslick at a given breakpoint now by adding:
    // settings: "unslick"
    // instead of a settings object
  ]
});
</script>