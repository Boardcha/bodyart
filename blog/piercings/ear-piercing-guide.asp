<%@LANGUAGE="VBSCRIPT"%>
<% 
Response.ContentType = "text/html"
Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8" 
%>
<%
	page_title = "Ear piercing location chart diagram"
	page_description = "A guide that will show you the names and locations of common ear piercings"
	page_keywords = ""
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<link rel=“canonical” href=“/blog/piercings/ear-piercing-guide.asp”  />
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<link rel="stylesheet" type="text/css" href="/CSS/slick.css"/>


    <div class="card  text-white bg-dark mt-3">
        <div class="card-header">
            <h3>Bodyartforms guide to ear piercings</h3>
        </div>
        <div class="card-body">
            Ever get confused about what the names of all the different types of ear piercings are? Look no further! We've created this easy to use guide where you'll learn about the the top ear piercings as well as what type of jewelry works best for them.
        </div>
      </div> 
    

    <div class="card bg-light mt-3">
        <div class="card-header">
          <h5>EAR PIERCING CHART DIAGRAM</h5>
        </div>
        <div class="card-body">
            <div class="col-xs-12 col-sm-12 col-md-6 col-lg-4 col-xl-3">
                <img class="img-fluid" src="/images/ear-diagram.png">
            </div>
        </div>
      </div>
FIND A WAY TO ALSO PULL IN CUSTOMER PHOTOS FROM THE GALLERIES JUST LIKE THE JEWELRY INCLUDES
      <div class="card bg-light mt-3">
        <div class="card-header">
          <h5>HELIX PIERCING</h5>
        </div>
        <div class="card-body">
            xxxx
            <a href="https://bodyartforms.com/products.asp?piercing=Helix" target="_blank">Here's a link to all our helix jewelry</a>

            <%
            SqlString = "SELECT TOP 20 * FROM FlatProducts WHERE tags LIKE '%helix%' AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC"
            %>
            <!--#include virtual="/includes/inc-embedded-products.inc" --> 
        </div>
      </div> 

      <div class="card bg-light mt-3">
        <div class="card-header">
          <h5>TRAGUS PIERCING</h5>
        </div>
        <div class="card-body">
            xxxx
            <a href="https://bodyartforms.com/products.asp?piercing=Tragus" target="_blank">Here's a link to all our tragus jewelry</a>

            <%
            SqlString = "SELECT TOP 20 * FROM FlatProducts WHERE tags LIKE '%tragus%' AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC"
            %>
            <!--#include virtual="/includes/inc-embedded-products.inc" --> 
        </div>
      </div> 




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