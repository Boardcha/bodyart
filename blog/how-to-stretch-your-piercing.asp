<%@LANGUAGE="VBSCRIPT"%>
<% 
Response.ContentType = "text/html"
Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8" 
%>
<%
	page_title = "How to stretch / gauge your ear piercing with a taper"
	page_description = "Learn how to stretch your ear piercing ... without having terrible things happen"
	page_keywords = ""
%>
<!--#include virtual="/functions/security.inc" -->
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->
<link rel="stylesheet" type="text/css" href="/CSS/slick.css"/>


    <div class="card  text-white bg-dark mt-3">
        <div class="card-header">
            <h3>How to stretch / gauge your ear piercing with a taper</h3>
            <h4>(Without having terrible things happen)</h4>
        </div>
        <div class="card-body">
            <p>You are about to embark on an incredible journey. Stretching a piercing is a practice that goes back thousands of years. The oldest mummified remains ever found, “Otzi” the caveman, had stretched lobes. Jewelry has been found in graves from the very first civilizations. It’s a practice that goes back so far that it may be as old as we are.</p>
            <p>And since the Neanderthals were known to wear jewelry as well, we might not even be the first humans to do it. Just think about that for a moment. Stretching lobes might be older than us.</p>
            <p>That having been said, this is a guide on how to stretch your ears, not the history of it (<a href="/blog/history-of-stretching.asp" target="_blank">That’s another blog</a>). </p>
            <p>All set? Good.</p>
        </div>
      </div> 
    

    <div class="card bg-light mt-3">
        <div class="card-header">
          <h5>STEP 1 - ONLY STRETCH A HEALED PIERCING</h5>
        </div>
        <div class="card-body">
            You’ll want your piercing to be 100% healed. If it's irritated, swollen, or pussing you'll need to take care of your piercing until it's fully healed and then you can proceed with stretching.
        </div>
      </div>


    <div class="card bg-light mt-3">
        <div class="card-header">
          <h5>STEP 2 - BE PATIENT</h5>
        </div>
        <div class="card-body">
            <p>Pay close attention to what might, and probably will, happen if you skip ahead and get impatient. Things like <a href="http://wiki.bme.com/index.php?title=Uneven_Stretching" target="_blank">uneven stretching</a>, <a href="http://wiki.bme.com/index.php?title=Piercing_Blow-out" target="_blank">blowouts</a> and <a href="http://wiki.bme.com/index.php?title=Earlobe_Tearing" target="_blank">tearing</a>.</p>
            <p>Right, so we want to be careful and always listen to our body and above all, be patient. Stretching is a process, it takes time. You might feel some discomfort, a bit of a burning sensation, but you shouldn’t feel much pain, and definitely shouldn’t bleed. If that kind of thing happens, stop.</p>
        </div>
      </div>

 
      <div class="card bg-light mt-3">
        <div class="card-header">
          <h5>STEP 3 - GET THE RIGHT JEWELRY & TOOLS</h5>
        </div>
        <div class="card-body">
            Get the right jewelry and tools. In the old days, folks used pens, paintbrush handles, whatever was handy. But these aren’t the old days anymore and you shouldn’t use those things. The best thing to use is a taper and a quality single flare or non flared plug that your newly stretched piercing can heal with. We have an assortment of <a href="https://bodyartforms.com/products.asp?jewelry=tapers" target="_blank">stretching tools here</a>.</p>

            <p>Note that the plugs have no flare, and they’re made of steel. Titanium or glass will work as well, and if you want to you can use a single flare plug, but not a double flare. The flare will basically stretch you farther than the taper and can lead to problems (See step two).</p>
            <p>So you have your taper (the pointy thing) and you have your jewelry (steel, titanium or glass) and we are ready for step 4</p>

            <%
            SqlString = "SELECT TOP 20 * FROM FlatProducts WHERE tags LIKE '%tapers%' AND tags NOT LIKE '%save%' AND picture <> 'nopic.gif' AND active = 1 AND customorder <> 'yes' ORDER BY qty_sold_last_7_days DESC, ProductID DESC"
            %>
            <!--#include virtual="/includes/inc-embedded-products.inc" -->  
            
        </div>
      </div>   


      <div class="card bg-light mt-3">
        <div class="card-header">
          <h5>STEP 4 - PREPARE YOUR PIERCING</h5>
        </div>
        <div class="card-body">
            Get wet. Take a shower. Or take a warm wet cloth and soak the piercing for a few minutes to ease up the tissue and help with the stretch. You should wash your hands, have some jojoba oil or vitamin e handy for lubricant.
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


      <div class="card  text-white bg-success mt-3">
        <div class="card-header">
          <h5>CONGRATULATIONS!</h5>
        </div>
        <div class="card-body">
            <p>Your stretch is done. Unless it started to hurt too bad, you felt a lot of resistance and had to stop, with a taper only partway in. In which case go ahead and take the taper out and put your old jewelry back in. You’ll have to wait a few weeks and give it another shot.</p>
            <p>If you stretched one lobe already, that’s okay too. Lots of people have one lobe that’s tighter than the other. Remember, we’re being patient (See Step Two)</p>
        </div>
      </div> 

    <div class="card bg-light mt-3">
        <div class="card-header">
          <h5>TREAT YOUR FRESHLY STRETCHED PIERCING LIKE A NEW PIERCING</h5>
        </div>
        <div class="card-body">
            A stretched piercing is a new piercing. You should treat it exactly like you would a new piercing until it is fully healed. If you’re wanting to go up to a larger gauge, you’ll need to wait a bit.
        </div>
      </div> 

      <div class="card bg-light mt-3">
        <div class="card-header">
          <h5>SO, YOU WANT TO GET BIGGER?</h5>
        </div>
        <div class="card-body">
            Here’s a handy dandy chart on how long you should wait between stretches at different gauges. It’s a guideline. If this chart and your lobes disagree, side with your lobes and give them some more rest.
            <ul>
                <li>Before your first stretch: Approximately 5 to 6 months</li>
                <li>16g to 14g - 1 month</li>
                <li>14g to 12g - 1 month</li>
                <li>12g to 10g - 1.5 months</li>
                <li>10g to 8g - 2 months</li>
                <li>8g to 6g - 3 months</li>
                <li>6g to 4g - 3 month</li>
                <li>4g to 2g - 3 months</li>
                <li>2g to 0g - 4 months</li>
                <li>0g to 00g - 4 months</li>
            </ul>
        </div>
      </div>   

    <div class="card  text-white bg-danger mt-3">
        <div class="card-header">
          <h5>THE POINT OF NO RETURN</h5>
        </div>
        <div class="card-body">
            <p>Oh, yeah. We need to talk about this. If you want to shrink your lobes back down to normal someday, you’ll have to take your plugs out and let them shrink up. You can do this just to go down a size as well. The common consensus is that 2g is the largest you can go and still shrink back to normal. Everyone is different, but that’s as close as you’ll get to a hard and fast number.</p>
            <p>Bigger than 2g and you’ll likely need a plastic surgeon to sew your ears up. So remember, you’re making a pretty permanent choice.</p>
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