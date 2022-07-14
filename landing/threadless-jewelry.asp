<%@LANGUAGE="VBSCRIPT"%>
<% 
Response.ContentType = "text/html"
Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8" 
%>
	<%
	page_title = "Threadless body jewelry"
	page_description = "Switch your body jewelry out quickly with quality threadless jewelry"
	page_keywords = "threadless body jewelry"
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<link rel="stylesheet" type="text/css" href="/CSS/slick.css"/>
<style>
	.img-circles img{border-width:3px!important}
	.img-circles img:hover{border-color: #696887!important;border-width:10px!important}
</style>
<!--#include virtual="/bootstrap-template/filters.asp" -->
<div class="slider-container text-center">
	<a href="/products.asp?threading=Threadless">
	<picture>
		<source media="(max-width: 550px)" sizes="100vw" srcset="/images/landing-pages/threadless/landing-threadless-jewelry-550x350.jpg">
		<source media="(max-width: 850px)" sizes="100vw" srcset="/images/landing-pages/threadless/landing-threadless-jewelry-850x350.jpg">
		<source media="(max-width: 1024px)" sizes="100vw" srcset="/images/landing-pages/threadless/landing-threadless-jewelry-1024x350.jpg">
		<source media="(max-width: 1600px)" sizes="100vw" srcset="/images/landing-pages/threadless/landing-threadless-jewelry-1600x350.jpg">
		<source media="(min-width: 1920px)" sizes="100vw" srcset="/images/landing-pages/threadless/landing-threadless-jewelry-1920x350.jpg">
			<img class="img-fluid" src="/images/landing-pages/threadless/landing-threadless-jewelry-1600x350.jpg" 
		srcset="/images/landing-pages/threadless/landing-threadless-jewelry-1920x350.jpg 1920w,
		/images/landing-pages/threadless/landing-threadless-jewelry-1600x350.jpg 1600w,
		/images/landing-pages/threadless/landing-threadless-jewelry-1024x350.jpg 1024w,
		/images/landing-pages/threadless/landing-threadless-jewelry-850x350.jpg 850w,
		/images/landing-pages/threadless/landing-threadless-jewelry-550x350.jpg 550w"
		sizes="100vw"
		alt="Threadless jewelry top main landing image" />
	</picture>
	</a>
</div>
<div class="p-2">
<h4 class="d-inline">How threadless body jewelry works</h4>
<div class="mb-4">
Threadless body jewelry is a great alternative to standard threaded jewelry. Threadless ends use a thin "pin" that you slightly bend to get a secure fit. Check out our video and images below to learn how threadless jewelry works.
</div>

<div class="container-fluid text-center">
	<div class="row justify-content-center">
		<div class="col-xl-6 col-lg-6 col-md-6 col-12">
			<video class="mw-100" width="560" height="315" preload="metadata" controls muted>
				<source src="https://videos.bodyartforms.com/video-threadless-ends-how-to.mp4#t=0.5" type="video/mp4">
			  Your browser does not support playing embedded videos
			  </video>
		</div>
		<div class="col-xl-6 col-lg-6 col-md-6 col-12">
			<div class="row justify-content-center">
			<div class="col-5 col-xl-5 col-lg-5 col-md-5 col-sm-5 p-2 mr-1 rounded border border-secondary">
				<img class="img-fluid" src="/images/landing-pages/threadless/threadless-how-to-1.png">
				<div>Insert the pin about 1/3 to halfway into the post and gently bend it downward.</div>
			</div>	
			<div class="col-5 col-xl-5 col-lg-5 col-md-5 col-sm-5 p-2 ml-1 rounded border border-secondary">
				<img class="img-fluid" src="/images/landing-pages/threadless/threadless-how-to-2.png">
				<div>Push the end all the way into the post to test the fit.<br>
					<strong>More bend = tighter fit</strong></div>
			</div>
		</div>
		</div>	
		</div>	
	</div><!-- row -->
</div><!-- container-->


<div class="container-fluid img-circles my-3">
	<div class="row  justify-content-center text-center">
		<div class="col-6 col-lg-3 col-xl-2 col-md-3 py-3">
			<a class="text-secondary" class="track-click" id="landing-threadless-balls-image" href="/products.asp?jewelry=balls&threading=Threadless">
				<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/threadless/threadless-balls-ends.jpg">
				<h4>THREADLESS BALLS<br>& ENDS</h4>	
			</a>
		</div>
		<div class="col-6 col-lg-3 col-xl-2 col-md-3 py-3">
			<a class="text-secondary" class="track-click" id="landing-threadless-labrets-image" href="/products.asp?jewelry=labret&threading=Threadless">
				<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/threadless/threadless-labrets.jpg">
				<h4>THREADLESS LABRETS</h4>	
			</a>
		</div>
		<div class="col-6 col-lg-3 col-xl-2 col-md-3 py-3">
			<a class="text-secondary" class="track-click" id="landing-threadless-nosescrews-image" href="/products.asp?jewelry=nose-ring&price=0%3B100&threading=Threadless">
				<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/threadless/threadless-nosescrews.jpg">
				<h4>THREADLESS NOSE<br>JEWELRY</h4>	
			</a>
		</div>
		<div class="col-6 col-lg-3 col-xl-2 col-md-3 py-3">
			<a class="text-secondary" class="track-click" id="landing-threadless-straight-barbells-image" href="/products.asp?jewelry=barbell&threading=Threadless">
				<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/threadless/threadless-straight-barbells.jpg">
				<h4>THREADLESS STRAIGHT<BR>BARBELLS</h4>	
			</a>
		</div>
		<div class="col-6 col-lg-3 col-xl-2 col-md-3 py-3">
			<a class="text-secondary" class="track-click" id="landing-threadless-curved-barbells-image" href="/products.asp?jewelry=curved&threading=Threadless">
				<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/threadless/threadless-curved-barbells.jpg">
				<h4>THREADLESS CURVED<BR>BARBELLS</h4>	
			</a>
		</div>
	</div>
</div><!-- container -->
<%
SqlString = "SELECT TOP 20 * FROM FlatProducts WHERE tags LIKE '%threadless%' AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC"
%>
<h4>Threadless Best Sellers</h4>
<!--#include virtual="/includes/inc-embedded-products.inc" -->
 
</div><!-- full page div -->

<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript" src="/js/slick-customized.min.js"></script>
<script type="text/javascript">
	$('.embedded-products').slick({
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