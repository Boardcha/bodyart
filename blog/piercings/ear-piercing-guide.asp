<%@LANGUAGE="VBSCRIPT"%>
<% 
Response.ContentType = "text/html"
Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8" 
%>
<%
	page_title = "Ear piercing guide and location chart diagram"
	page_description = "A guide that will show you the names and locations of the common ear piercings"
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
<style>
  .slick-track {
      margin-left:0;
  }
</style>


    <div class="card  text-white bg-dark mt-3">
        <div class="card-header">
            <h3>Bodyartforms guide to ear piercings</h3>
        </div>
        <div class="card-body">
            <p>Ever get confused about what the names of all the different types of ear piercings are? Look no further! We've created this easy-to-use guide where you'll learn about the top ear piercings as well as what type of jewelry works best for them.</p>

            <h5>JUMP TO:</h5>
            <ul>
              <li><a class="text-info" href="#helix">Helix</a></li>
              <li><a class="text-info" href="#tragus">Tragus</a></li>
              <li><a class="text-info" href="#daith">Daith</a></li>
              <li><a class="text-info" href="#industrial">Industrial</a></li>
              <li><a class="text-info" href="#conch">Conch</a></li>
              <li><a class="text-info" href="#rook">Rook</a></li>
              <li><a class="text-info" href="#antitragus">Anti Tragus</a></li>
              <li><a class="text-info" href="#snug">Snug</a></li>
            </ul>
        </div>
      </div> 
    
      <div class="card bg-light mt-3">
        <div class="card-header">
          <h3 id="helix">HELIX PIERCING</h3>
        </div>
        <div class="card-body">
            <div class="row">
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-4 col-xl-3">
                    <img class="img-fluid" src="/images/blog/piercings/helix-piercing.png" />
                </div>
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-8 col-xl-9">
                    A helix piercing is any piercing on the upper cartilage of your ear. There are actually different types of helix piercings: standard single helix (one piercing), double helix (two piercings), triple helix (three piercings), forward helix, and the <a href="#snug">anti-helix (or snug)</a>.
                    <br>
                    <a class="btn btn-sm btn-purple my-2" href="/products.asp?piercing=Helix" target="_blank">SHOP OUR HELIX JEWELRY</a>

                    <%
                    SqlString = "SELECT TOP 3 * FROM FlatProducts WHERE tags LIKE '%helix%' AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC"
                    %>
                    <!--#include virtual="/includes/inc-embedded-products.inc" -->
                </div>
            </div> 
        </div>
      </div> 

      <div class="card bg-light mt-3">
        <div class="card-header">
          <h3 id="tragus">TRAGUS PIERCING</h3>
        </div>
        <div class="card-body">
            <div class="row">
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-4 col-xl-3">
                    <img class="img-fluid"  src="/images/blog/piercings/tragus-piercing.png" />
                </div>
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-8 col-xl-9">
                    <p>
                        The tragus piercing is located on the small area of cartilage right next to your ear canal. Even though the location is small, you can still wear <a href="/products.asp?jewelry=labret&piercing=Tragus" target="_blank">studs</a>, <a href="/products.asp?jewelry=captive&piercing=Tragus" target="_blank">rings</a>, and <a href="/products.asp?jewelry=barbell&piercing=Tragus" target="_blank">barbells</a> in a tragus piercing.
                    </p>
                    <a class="btn btn-sm btn-purple" href="/products.asp?piercing=Tragus" target="_blank">SHOP OUR TRAGUS JEWELRY</a>

                    <%
                    SqlString = "SELECT TOP 3 * FROM FlatProducts WHERE tags LIKE '%tragus%' AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC"
                    %>
                    <!--#include virtual="/includes/inc-embedded-products.inc" --> 
                </div>
            </div>
            



        </div>
      </div> 

      <div class="card bg-light mt-3">
        <div class="card-header">
          <h3 id="daith">DAITH PIERCING</h3>
        </div>
        <div class="card-body">
            <div class="row">
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-4 col-xl-3">
                    <img class="img-fluid" src="/images/blog/piercings/diath-piercing.png" />
                </div>
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-8 col-xl-9">
                    <p>The daith piercing is located on the inner cartilage of the ear right near the conch area. <a href="/products.asp?jewelry=captive&piercing=Daith" target="_blank">Rings</a> and <a href="/products.asp?jewelry=curved&piercing=Daith" target="_blank">curved</a> jewelry are the easiest to wear in the daith piercing.
                    </p>
                    <a class="btn btn-sm btn-purple" href="/products.asp?piercing=Daith" target="_blank">SHOP OUR DAITH JEWELRY</a>

                    <%
                    SqlString = "SELECT TOP 3 * FROM FlatProducts WHERE tags LIKE '%daith%' AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC"
                    %>
                    <!--#include virtual="/includes/inc-embedded-products.inc" --> 
                </div>
            </div>
        </div>
      </div> 

      <div class="card bg-light mt-3">
        <div class="card-header">
          <h3 id="industrial">INDUSTRIAL PIERCING</h3>
        </div>
        <div class="card-body">
            <div class="row">
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-4 col-xl-3">
                    <img class="img-fluid" src="/images/blog/piercings/industrial-piercing.png" />
                </div>
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-8 col-xl-9">
                    <p>
                    An industrial piercing is two separate ear cartilage piercings that typically use one long bar to connect them. It's always important to discuss with your piercer whether an industrial piercing will be appropriate for your ear as everyone's anatomy is different.
                    </p>
                    The majority of industrial jewelry is long straight barbells. There are also some neat designs where it's two separate bars attached via chains. Another option to switch your style up is to skip the bar all together and use rings or studs in your piercings.
                    <br>
                    <a class="btn btn-sm btn-purple my-2" href="/products.asp?piercing=Industrial" target="_blank">SHOP OUR INDUSTRIAL JEWELRY</a>

                    <%
                    SqlString = "SELECT TOP 3 * FROM FlatProducts WHERE tags LIKE '%industrial%' AND tags NOT LIKE '%save%' AND tags NOT LIKE '%aftercare%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC"
                    %>
                    <!--#include virtual="/includes/inc-embedded-products.inc" --> 
                </div>
            </div>

        </div>
      </div> 

      <div class="card bg-light mt-3">
        <div class="card-header">
          <h3 id="conch">CONCH PIERCING</h3>
        </div>
        <div class="card-body">
            <div class="row">
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-4 col-xl-3">
                    <img class="img-fluid" src="/images/blog/piercings/conch-piercing.png" />
                </div>
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-8 col-xl-9">
                    <p>
                        The conch piercing is a single piercing through the center cartilage of the ear. There are two styles: inner conch and outer conch. 
                    </p>
                        The outer conch is pierced close to the rim of the ear and typically is adorned with <a href="/products.asp?jewelry=captive&piercing=Conch" target="_blank">rings</a>. The inner conch is pierced further from the rim of the ear and is best suited for <a href="/products.asp?jewelry=labret&piercing=Conch" target="_blank">studs</a> although large diameter <a href="/products.asp?jewelry=captive&piercing=Conch&length=5%2F8"&length=3%2F4"&length=7%2F8"&length=1"" target="_blank">rings</a> can be worn as well.

                    <br>
                    <a class="btn btn-sm btn-purple my-2" href="/products.asp?piercing=Conch" target="_blank">SHOP OUR CONCH JEWELRY</a>

                    <%
                    SqlString = "SELECT TOP 3 * FROM FlatProducts WHERE tags LIKE '%conch%' AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC"
                    %>
                    <!--#include virtual="/includes/inc-embedded-products.inc" --> 
                </div>
            </div>

        </div>
      </div> 

      <div class="card bg-light mt-3">
        <div class="card-header">
          <h3 id="rook">ROOK PIERCING</h3>
        </div>
        <div class="card-body">
            <div class="row">
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-4 col-xl-3">
                    <img class="img-fluid" src="/images/blog/piercings/rook-piercing.png" />
                </div>
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-8 col-xl-9">
                    <p>The rook piercing is a cartilage piercing in the upper cartilage of your ear, located above the tragus.
                    </p>
                    <a class="btn btn-sm btn-purple" href="/products.asp?piercing=Rook" target="_blank">SHOP OUR ROOK JEWELRY</a>

                    <%
                    SqlString = "SELECT TOP 3 * FROM FlatProducts WHERE tags LIKE '%rook%' AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC"
                    %>
                    <!--#include virtual="/includes/inc-embedded-products.inc" --> 
                </div>
            </div>
        </div>
      </div> 

      
      <div class="card bg-light mt-3">
        <div class="card-header">
          <h3 id="antitragus">ANTI TRAGUS PIERCING</h3>
        </div>
        <div class="card-body">
            <div class="row">
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-4 col-xl-3">
                    <img class="img-fluid" src="/images/blog/piercings/anti-tragus-piercing.png" />
                </div>
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-8 col-xl-9">
                    <p>The anti tragus piercing is located directly across from the tragus piercing on the rim of cartilage above your earlobe.
                    </p>
                    <a class="btn btn-sm btn-purple" href="/products.asp?piercing=Anti Tragus" target="_blank">SHOP OUR ANTI TRAGUS JEWELRY</a>

                    <%
                    SqlString = "SELECT TOP 3 * FROM FlatProducts WHERE tags LIKE '%antitragus%' AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC"
                    %>
                    <!--#include virtual="/includes/inc-embedded-products.inc" --> 
                </div>
            </div>

        </div>
      </div> 

      <div class="card bg-light mt-3">
        <div class="card-header">
          <h3 id="snug">SNUG PIERCING</h3>
        </div>
        <div class="card-body">
            <div class="row">
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-4 col-xl-3">
                    <img class="img-fluid" src="/images/blog/piercings/snug-piercing.png" />
                </div>
                <div class="col-xs-12 col-sm-12 col-md-6 col-lg-8 col-xl-9">
                    <P>The snug is nicknamed the "anti-helix" piercing. It's located in between the rim of your ear and your inner cartilage, and above the anti-tragus. 
                    </P>
                    <a class="btn btn-sm btn-purple" href="/products.asp?piercing=Snug" target="_blank">SHOP OUR SNUG JEWELRY</a>

                    <%
                    SqlString = "SELECT TOP 3 * FROM FlatProducts WHERE tags LIKE '%snug%' AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC"
                    %>
                    <!--#include virtual="/includes/inc-embedded-products.inc" --> 
                </div>
            </div>

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