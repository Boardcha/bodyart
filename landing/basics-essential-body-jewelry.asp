<%@LANGUAGE="VBSCRIPT"%>
	<%
	page_title = "Essential body jewelry - Rings, loose ends, barbells, labrets, curves, and circulars"
	page_description = "We have all the essential body jewelry you need for your piercing at very reasonable prices. Sister owned since 2001!"
	page_keywords = "captive rings, essential body jewelry"
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<style>
	.img-circles img{border-width:3px!important}
	.img-circles img:hover{border-color: #696887!important;border-width:10px!important}
</style>

<div class="slider-container text-center">
	<picture>
		<source media="(max-width: 550px)" sizes="100vw" srcset="/images/landing-pages/basics/landing-basics-jewelry-550x350.jpg">
		<source media="(max-width: 850px)" sizes="100vw" srcset="/images/landing-pages/basics/landing-basics-jewelry-850x350.jpg">
		<source media="(max-width: 1024px)" sizes="100vw" srcset="/images/landing-pages/basics/landing-basics-jewelry-1024x350.jpg">
		<source media="(max-width: 1600px)" sizes="100vw" srcset="/images/landing-pages/basics/landing-basics-jewelry-1600x350.jpg">
		<source media="(min-width: 1920px)" sizes="100vw" srcset="/images/landing-pages/basics/landing-basics-jewelry-1920x350.jpg">
			<img class="img-fluid" src="/images/landing-pages/basics/landing-basics-jewelry-1600x350.jpg" 
		srcset="/images/landing-pages/basics/landing-basics-jewelry-1920x350.jpg 1920w,
		/images/landing-pages/basics/landing-basics-jewelry-1600x350.jpg 1600w,
		/images/landing-pages/basics/landing-basics-jewelry-1024x350.jpg 1024w,
		/images/landing-pages/basics/landing-basics-jewelry-850x350.jpg 850w,
		/images/landing-pages/basics/landing-basics-jewelry-550x350.jpg 550w"
		sizes="100vw"
		alt="Regular jewelry top main landing image" />
	</picture>
</div>

<div class="container-fluid img-circles">
	<div class="row  justify-content-center text-center">
		<div class="col-6 col-lg-3 col-xl-3 col-md-3 py-3">
			<a class="text-secondary" class="track-click" id="landing-regular-rings-image" href="/products.asp?jewelry=captive">
				<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/basics/landing-rings.jpg">
				<h4>RINGS</h4>	
			</a>
				
		</div>
		<div class="col-6 col-lg-3 col-xl-3 col-md-3 py-3">
			<a class="text-secondary" class="track-click" id="landing-regular-rings-image" href="/products.asp?jewelry=balls">
				<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/basics/landing-ends.jpg">
				<h4>LOOSE ENDS</h4>	
			</a>
				
		</div>
		<div class="col-6 col-lg-3 col-xl-3 col-md-3 py-3">
			<a class="text-secondary" href="/products.asp?jewelry=barbell">
				<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/basics/landing-straight-barbells.jpg">
				<h4>STRAIGHT BARBELLS</h4>
			</a>
			
		</div>
		<div class="col-6 col-lg-3 col-xl-3 col-md-3 py-3">
			<a class="text-secondary" href="/products.asp?jewelry=labret">
				<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/basics/landing-labrets.jpg">
				<h4>LABRETS</h4>
			</a>
			
		</div>
		<div class="col-6 col-lg-3 col-xl-3 col-md-3 py-3">
			<a class="text-secondary" href="/products.asp?jewelry=curved">
			<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/basics/landing-curved-barbells.jpg">
			<h4>CURVED BARBELLS</h4>
		</a>
		</div>
        <div class="col-6 col-lg-3 col-xl-3 col-md-3 py-3">
			<a class="text-secondary" href="/products.asp?jewelry=circular">
				<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/basics/landing-circular-barbells.jpg">
				<h4>CIRCULAR BARBELLS</h4>
			</a>
			
		</div>
	</div>


</div><!-- container -->


<!--#include virtual="/bootstrap-template/footer.asp" -->