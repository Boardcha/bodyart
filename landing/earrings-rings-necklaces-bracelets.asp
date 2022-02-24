<%@LANGUAGE="VBSCRIPT"%>
	<%
	page_title = "Rings, necklaces, chains, bracelets & more"
	page_keywords = "regular finger rings, necklaces, necklace chains, bracelets"
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->


<div class="slider-container text-center">
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
</div>
LINK MAIN BANNER TO ALL products

<div class="container-fluid">
	<div class="row text-center">
		<div class="col-6 col-lg-3 col-xl-2 py-3">
			<a class="track-click" id="landing-regular-rings-image" href="/products.asp?jewelry=finger-ring">
				<img class="img-fluid border border-secondary rounded-circle" style="border-width: 3px !important;" src="/images/landing-pages/regular-jewelry/rings.jpg">
				</a>
			
				<h4>RINGS</h4>
			
			
		</div>
		<div class="col-6 col-lg-3 col-xl-2 py-3">
			<a href="/products.asp?jewelry=necklace">
				<img class="img-fluid border border-secondary rounded-circle" style="border-width: 3px !important;" src="/images/landing-pages/regular-jewelry/necklaces.jpg">
			</a>
			<h4>NECKLACES</h4>
			<a class="d-block" href="/products.asp?jewelry=chains-necklace">Necklace chains</a>
			Pendants
		</div>
		<div class="col-6 col-lg-3 col-xl-2 py-3">
			<a href="/products.asp?jewelry=bracelet">
				<img class="img-fluid border border-secondary rounded-circle" style="border-width: 3px !important;" src="/images/landing-pages/regular-jewelry/bracelets.jpg">
			</a>
			<h4>BRACELETS</h4>
		</div>
		<div class="col-6 col-lg-3 col-xl-2 py-3">
			<img class="img-fluid border border-secondary rounded-circle" style="border-width: 3px !important;" src="/images/landing-pages/regular-jewelry/ear-cuffs.jpg">
			<h4>EAR CUFFS</h4>
		</div>
	</div>


</div><!-- container -->


<!--#include virtual="/bootstrap-template/footer.asp" -->