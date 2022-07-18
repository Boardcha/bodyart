<%@LANGUAGE="VBSCRIPT"%>
<% 
Response.ContentType = "text/html"
Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8" 
%>
<%
	page_title = "Gauges / Plugs Sizing Conversion Measurement Guide"
	page_description = "Gauge jewelry conversion and sizing chart for stretched ears"
	page_keywords = ""
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<link rel=“canonical” href=“/blog/help/gauges-plugs-conversion-guide.asp”  />
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<link rel="stylesheet" type="text/css" href="/CSS/slick.css"/>


    <div class="card  text-white bg-dark mt-3">
        <div class="card-header">
            <h3>Gauges / Plugs Sizing Conversion Measurement Guide</h3>
        </div>
        <div class="card-body">
            Once you start stretching your ears (or any piercing) it becomes more important to know the conversions from inches to millimeters that way you can be sure you order the right jewelry and also the proper tools you'll need to stretch up to the next size.
        </div>
      </div> 
    

    <div class="card bg-light mt-3">
        <div class="card-header">
          <h5>GAUGE / PLUG CONVERSIONS</h5>
        </div>
        <div class="card-body">
          <table class="table w-auto">
            <thead class="thead-dark">
              <tr>
                <th class="table-cell">Gauge</th>
                <th class="table-cell">Millimeter</th>
                <th class="table-cell">Inch</th>
              </tr>
            </thead>
            <tr>
              <td>20g</td>
              <td>.81mm</td>
              <td></td>
            </tr>
            <tr>
              <td>18g</td>
              <td>1mm</td>
              <td></td>
            </tr>
              <tr>
                <td>16g</td>
                <td>1.2mm</td>
                <td>3/64&quot;</td>
              </tr>
              <tr>
                <td>14g</td>
                <td>1.6mm</td>
                <td>1/16&quot;</td>
              </tr>
              <tr>
                <td>12g</td>
                <td>2mm</td>
                <td>5/64&quot;</td>
              </tr>
              <tr>
                <td>10g</td>
                <td>2.4mm</td>
                <td>3/32&quot;</td>
              </tr>
              <tr>
                <td>9g</td>
                <td>2.7mm</td>
                <td></td>
              </tr>
              <tr>
                <td>8g</td>
                <td>3.2mm</td>
                <td>1/8&quot;</td>
              </tr>
              <tr>
                <td>7g</td>
                <td>3.5mm</td>
                <td></td>
              </tr>
              <tr>
                <td>6g</td>
                <td>4mm</td>
                <td>5/32&quot;</td>
              </tr>
              <tr>
                <td>4g</td>
                <td>5mm</td>
                <td>3/16&quot;</td>
              </tr>
              <tr>
                <td>5g</td>
                <td>4.5mm</td>
                <td></td>
              </tr>
              <tr>
                <td>2g</td>
                <td>6mm</td>
                <td>1/4&quot;</td>
              </tr>
              <tr>
                <td>1g</td>
                <td>7mm</td>
                <td></td>
              </tr>
              <tr>
                <td>0g</td>
                <td>8mm</td>
                <td>5/16&quot;</td>
              </tr>
              <tr>
                <td>00g</td>
                <td>9mm to 10mm</td>
                <td>3/8&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>11mm</td>
                <td>7/16&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>13mm</td>
                <td>1/2&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>14mm</td>
                <td>9/16&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>16mm</td>
                <td>5/8&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>19mm</td>
                <td>3/4&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>22mm</td>
                <td>7/8&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>25mm</td>
                <td>1&quot;</td>
              </tr>
              <tr>
                <td class="text-center" colspan="3">
                  <span class="btn btn-sm btn-outline-secondary" id="toggle-above-1inch">
                    Show sizes above 1 inch
                  </span>
                </td>
              </tr>
              <tbody style="display:none" id="above-1inch">
              <tr>
                <td>&nbsp;</td>
                <td>26.9mm</td>
                <td>1-1/16&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>28.5mm</td>
                <td>1-1/8&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>30.1mm</td>
                <td>1-3/16&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>31.7mm</td>
                <td>1-1/4&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>33.3mm</td>
                <td>1-5/16&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>34.9mm</td>
                <td>1-3/8&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>36.5mm</td>
                <td>1-7/16&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>38.1mm</td>
                <td>1-1/2&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>39.6mm</td>
                <td>1-9/16&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>41.2mm</td>
                <td>1-5/8&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>44.4mm</td>
                <td>1-3/4&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>47.6mm</td>
                <td>1-7/8&quot;</td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>51mm</td>
                <td>2&quot;</td>
              </tr>
            </tbody>
            </table>
        </div>
      </div>




 




      <div class="card bg-light mt-3">
        <div class="card-header">
          <h5>STEP 5 - LUBRICATION & THE STRETCH</h5>
        </div>
        <div class="card-body">
            <p>Apply some of your lube to the taper, particularly the pointy end, and go about halfway up. Don’t use so much oil that the taper is hard to hang on to. Place the pointy end of the taper into your piercing and begin to insert it. Give it a little twist before you get too far to spread the lubricant around a bit. Push the taper in until you feel resistance and check yourself.</p>
            <p>Is the taper more than halfway through? If so, you’re probably all set. Continue to slowly push until you’re at the level end of the taper. Line up your jewelry with the end of the taper (The gauge should be the exact same) and slide the rest of the way. You should now have a plug in your ear and the taper out the back.</p>
            <a href="https://bodyartforms.com/products.asp?jewelry=lotion-oil" target="_blank">Here's a link to all our piercing friendly oils and salves</a>

            <%
            SqlString = "SELECT TOP 20 * FROM FlatProducts WHERE tags LIKE '%lotionoil%' AND (title LIKE '%stretch%' OR title LIKE '%oil%') AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC"
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