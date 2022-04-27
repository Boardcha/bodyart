</div><!-- divs to end col layout for filters and main  -->
</div>
</div>
</div>
</main><!-- end main content -->


<footer>
    <div class="container-fluid text-light pb-2 pt-4 mt-5 px-md-4">
        <div class="row">
            <div class="col-6 col-lg-3 px-md-3">
                <div class="border-bottom pb-1 h5">
                        SUPPORT
                </div>
                <ul class="list-unstyled pt-2">
                    <li class="py-1 py-md-0">
                        <a href="/contact.asp" class="text-light">
                            Contact Us
                        </a>
                    </li>
                    <li class="py-1 py-md-0">
                            <a class="text-light" href="/faqs.asp">Help / FAQs</a>
                        </li>
                        <li class="py-1 py-md-0">
                            <a class="text-light" href="/help_MeasurementGuide.asp">Sizing & Measurements</a>
                        </li>
                        <li class="text-capitalize py-1 py-md-0 d-md-block">
                            <a class="text-light" href="/returns.asp">Return policy</a>
                        </li>
                        <li class="text-capitalize py-1 py-md-0 d-md-block">
                            <a class="text-light" href="/about.asp">About our team</a>
                        </li>
                </ul>
                <% if request.cookies("darkmode") <> "on" then
                    darkchecked = "" 
                else 
                    darkchecked = "checked"
                end if %>
                <span class="ml-1 font-weight-bold">DARK MODE</span>
                <div class="onoffswitch mb-5">
                    <input type="checkbox" name="onoffswitch" class="onoffswitch-checkbox" id="darkmode-switch" <%= darkchecked %>>
                    <label class="onoffswitch-label" for="darkmode-switch">
                        <span class="onoffswitch-inner"></span>
                        <span class="onoffswitch-switch"></span>
                    </label>
                </div>
            </div>
            <div class="col-6 col-lg-3 px-md-3">
                <div class="border-bottom pb-1 h5">
                    ACCOUNT
                </div>
                <ul class="list-unstyled pt-2">
                    <li class="py-1 py-md-0">
                        <a class="text-light" href="/sign-in.asp" data-toggle="modal" data-target="#signin" href="#">Sign In</a>
                    </li>
                    <% If not rsGetUser.EOF and request.cookies("ID") <> "" then %>
                    <li class="text-capitalize py-1 py-md-0">
                        <a class="text-light" href="/account.asp">Manage Account</a>
                    </li>
                    <li class="text-capitalize py-1 py-md-0">
                        <a class="text-light" href="/wishlist.asp">Wishlist</a>
                    </li>
                    <li class="text-capitalize py-1 py-md-0">
                        <a class="text-light" href="/sign-out.asp">Sign Out</a>
                    </li>
                    <% end if %>
                    <li class="text-capitalize py-1 py-md-0">
                        <a class="text-light" data-toggle="modal" data-target="#createaccount" href="#">Create Account</a>
                    </li>
                </ul>
            </div>
            <div class="col-6 col-lg-3 px-md-3">
                <div class="border-bottom pb-1 h5">
                    SHOP
                </div>
                <ul class="list-unstyled pt-2">
                    <li class="text-capitalize py-1 py-md-0">
                        <a class="text-light" href="/products.asp?new=Yes">New jewelry</a>
                    </li>
                    <li class="text-capitalize py-1 py-md-0">
                        <a class="text-light" href="/products.asp?feature=top_seller">Top sellers</a>
                    </li>
                    <li class="text-capitalize py-1 py-md-0">
                        <a class="text-light" href="/products.asp?jewelry=captive">Rings</a>
                    </li>
                    <li class="text-capitalize py-1 py-md-0">
                        <a class="text-light" href="/products.asp?jewelry=septum">Septum jewelry</a>
                    </li>
                    <li class="text-capitalize py-1 py-md-0">
                        <a class="text-light" href="/products.asp?jewelry=nose-ring">Nose jewelry</a>
                    </li>
                    <li class="text-capitalize py-1 py-md-0">
                        <a class="text-light" href="/products.asp?jewelry=plugs">Plugs & stretched piercings</a>
                    </li>
                    <li class="text-capitalize py-1 py-md-0">
                        <a class="text-light" href="/products.asp?jewelry=balls">Balls, ends, & beads</a>
                    </li>
                    <li class="text-capitalize py-1 py-md-0">
                        <a class="text-light" href="/landing/earrings-rings-necklaces-bracelets.asp">Earrings, rings, necklaces, & bracelets</a>
                    </li>
                </ul>
            </div>
            <div class="col-12 col-md-6 col-lg-3 px-md-3">
                <div class="border-bottom pb-1 h5">
                    STAY CONNECTED
                </div>
                <div class="py-2">Sign up for our newsletter and get notified anytime we run sales or special events</div>
                <form name="ccoptin" target="_blank" class="">
                    <div class="form-group mb-2">
                        <input class="form-control" placeholder="E-mail address" type="text" name="footer_newsletter_email" id="footer_newsletter_email" />
                    </div>
                    <span class="btn btn-purple event-newsletter" id="footer-newsletter-signup">Sign Up!</span><span  id="footer-newsletter-msg"></span>
                </form>
                <div class="py-3">
                    <a href="http://instagram.com/bodyartforms" target="_blank">
                        <i class="fa fa-instagram fa-2x text-light pr-1"></i>
                    </a>
                    <a href="https://www.facebook.com/pages/Bodyartforms/149344708430326" target="_blank">
                        <i class="fa fa-facebook-square fa-2x text-light px-1"></i>
                    </a>
                    <a href="https://www.pinterest.com/bodyartforms/" target="_blank">
                        <i class="fa fa-pinterest-square fa-2x text-light px-1"></i>
                    </a>
                    <a href="https://twitter.com/bodyartforms" target="_blank">
                        <i class="fa fa-twitter-square fa-2x text-light px-1"></i>
                    </a>
                    <a class="btn btn-sm btn-light" href="https://g.page/r/CVDk_0MEUfIlEAQ/review" target="_blank" >
                        <img src="/images/homepage/google-icon.png" class="mr-1" style="height: 20px" /> Review us on Google
                    </a>
                </div>
            </div>
        </div>
        <div class="row py-2">
            <div class="text-center w-100">
                <a href="/privacy-policy.asp" class="text-info">
                    Privacy Policy</a>
                &copy;
                <%= year(date) %> Bodyartforms LLC
                <br/>
                <img class="my-2" src="/images/ssl-secure-site.png" />
            </div>
        </div>
    </div>
</footer>
<script src="/js/jquery-3.3.1.min.js"></script>
<script src="/js/popper.min.js"></script>
<script src="/js/bootstrap-v4.min.js"></script>
<script type="text/javascript" src="/js/js.cookie.js"></script>
<script type="text/javascript" src="/js/lazysizes.min.js"></script>
<!--Range slider Plugin files and settings-->
<link rel="stylesheet" href="/css/ion.rangeSlider.min.css"/>
<script src="/js/ion.rangeSlider.min.js"></script>
<script type="text/javascript">
	$("#price-range").ionRangeSlider({
        skin: "round",
		max_postfix : "+",
		prefix : "$",
		values_separator : " - "
    });

</script> 

</script> 

<!--Google Sign-in-->
<script src="https://apis.google.com/js/platform.js?onload=init" async defer></script>
<script>
   function init() {
        gapi.load('auth2', function() {
            auth2 = gapi.auth2.init({
                client_id: '534605611929-1n69ud9b72jvjh21okl82g4sjttdjmj8.apps.googleusercontent.com',
                cookiepolicy: 'single_host_origin',
                scope: 'profile email'
            });
			renderButton();
            element = document.getElementById('google_sign_in');
            auth2.attachClickHandler(element, {}, onSignIn, onFailure);
        });
    }
    function onSignIn(googleUser) {
		var profile = googleUser.getBasicProfile();	
		var id_token = googleUser.getAuthResponse().id_token;
		var xhr = new XMLHttpRequest();
		xhr.open('POST', '/oauth/google-signin-token.asp');
		xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
		xhr.overrideMimeType("application/json");
		xhr.onload = function() {
		  var json = JSON.parse(xhr.responseText);
		  if(json.status == "logged-in")
		    googleUser.disconnect();
			window.location.href = "/account.asp";
		};
		xhr.send("idtoken=" + id_token);
		
	}
    function onFailure(error) {
      console.log(error);
    }
    function renderButton() {
		gapi.signin2.render('google_sign_in', {
		'scope': 'profile email',
		'width': 240,
		'height': 50,
		'longtitle': true,
		'theme': 'dark'
		});
    }
</script><!--End Google Sign-in-->


<% ' ---- ONLY SHOW TO COUNTRIES IN THE EU -----------
%>
<!--#include virtual="/functions/inc-eu-country-codes.asp" -->
<script type="text/javascript" id="cookieinfo"
        // credit goes to https://cookieinfoscript.com
        src="/js/cookienotice.min.js"
        data-bg="#645862"
        data-fg="#FFFFFF"
        data-link="#F1D600"
        data-cookie="CookieInfoScript"
        data-text-align="left"
        data-message="We use cookies to give you the best possible user experience. Read our cookie policy <a href='privacy-policy.asp#cookies'>here</a> to learn more about our use of cookies and how to change your browser settings. By continuing to use this site you agree to the use of cookies."
        data-scriptmsg = ""
        data-moreinfo = "/privacy-policy.asp#cookies"
        data-close-text="Got it!">
    </script>
<% end if ' ONLY TO EU COUNTRIES %>
 <script type="text/javascript" src="/js-pages/footer.min.js?v=02172022_v2"></script>  
<% if request.cookies("adminuser") = "yes" then %>
<script>
    // Toggle sandbox front end load
    $(".toggle-sandbox").on('click',function(){
            
        var toggle_status = $(this).attr("data-sandbox");
        $.ajax({
              url: "sandbox.asp",
              data: { sandbox: toggle_status }
            }).done(function() {
              location.reload();
            });		
    });

    </script>
  <% end if ' logged in as admin user %>
<!-- Facebook Pixel Code -->
<script>
        !function(f,b,e,v,n,t,s){if(f.fbq)return;n=f.fbq=function(){n.callMethod?
        n.callMethod.apply(n,arguments):n.queue.push(arguments)};if(!f._fbq)f._fbq=n;
        n.push=n;n.loaded=!0;n.version='2.0';n.queue=[];t=b.createElement(e);t.async=!0;
        t.src=v;s=b.getElementsByTagName(e)[0];s.parentNode.insertBefore(t,s)}(window,
        document,'script','https://connect.facebook.net/en_US/fbevents.js');
        
        fbq('init', '532347420293260');
        fbq('track', "PageView");</script>
        <noscript><img height="1" width="1" style="display:none"
        src="https://www.facebook.com/tr?id=532347420293260&ev=PageView&noscript=1"
        /></noscript>
        <!-- End Facebook Pixel Code -->
</body>
</html>
<%
DataConn.Close()
Set DataConn = Nothing
%>