<%@LANGUAGE="VBSCRIPT"%>
	<%
	page_title = "Unique rings, necklaces, pendants, bracelets & ear cuffs"
	page_description = "Shop our selection of necklaces, pendants, rings, ear cuffs, and bracelets. Express yourself."
	page_keywords = "regular finger rings, necklaces, necklace chains, bracelets"
%>
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
	<a href="/products.asp?jewelry=bracelet&jewelry=earring&jewelry=necklace&jewelry=finger-ring">
	<picture>
		<source media="(max-width: 550px)" sizes="100vw" srcset="/images/landing-pages/regular-jewelry/landing-regular-jewelry-550x350.jpg">
		<source media="(max-width: 850px)" sizes="100vw" srcset="/images/landing-pages/regular-jewelry/landing-regular-jewelry-850x350.jpg">
		<source media="(max-width: 1024px)" sizes="100vw" srcset="/images/landing-pages/regular-jewelry/landing-regular-jewelry-1024x350.jpg">
		<source media="(max-width: 1600px)" sizes="100vw" srcset="/images/landing-pages/regular-jewelry/landing-regular-jewelry-1600x350.jpg">
		<source media="(min-width: 1920px)" sizes="100vw" srcset="/images/landing-pages/regular-jewelry/landing-regular-jewelry-1920x350.jpg">
			<img class="img-fluid" src="/images/landing-pages/regular-jewelry/landing-regular-jewelry-1600x350.jpg" 
		srcset="/images/landing-pages/regular-jewelry/landing-regular-jewelry-1920x350.jpg 1920w,
		/images/landing-pages/regular-jewelry/landing-regular-jewelry-1600x350.jpg 1600w,
		/images/landing-pages/regular-jewelry/landing-regular-jewelry-1024x350.jpg 1024w,
		/images/landing-pages/regular-jewelry/landing-regular-jewelry-850x350.jpg 850w,
		/images/landing-pages/regular-jewelry/landing-regular-jewelry-550x350.jpg 550w"
		sizes="100vw"
		alt="Regular jewelry top main landing image" />
	</picture>
	</a>
</div>

<div class="container-fluid img-circles">
	<div class="row  justify-content-center text-center">
		<div class="col-6 col-lg-3 col-xl-3 col-md-3 py-3">
			<a class="text-secondary" class="track-click" id="landing-regular-rings-image" href="/products.asp?jewelry=finger-ring">
				<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/regular-jewelry/rings.jpg">
				<h4>RINGS</h4>	
			</a>
				
		</div>
		<div class="col-6 col-lg-3 col-xl-3 col-md-3 py-3">
			<a class="text-secondary" href="/products.asp?jewelry=necklace">
				<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/regular-jewelry/necklaces.jpg">
				<h4>NECKLACES</h4>
			</a>
			
		</div>
		<div class="col-6 col-lg-3 col-xl-3 col-md-3 py-3">
			<a class="text-secondary" href="/products.asp?jewelry=bracelet">
				<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/regular-jewelry/bracelets.jpg">
				<h4>BRACELETS</h4>
			</a>
			
		</div>
		<div class="col-6 col-lg-3 col-xl-3 col-md-3 py-3">
			<a class="text-secondary" href="/products.asp?keywords=ear+cuff">
			<img class="img-fluid border border-secondary rounded-circle" src="/images/landing-pages/regular-jewelry/ear-cuffs.jpg">
			<h4>EAR CUFFS</h4>
		</a>
		</div>
	</div>


</div><!-- container -->


<!--#include virtual="/bootstrap-template/footer.asp" -->