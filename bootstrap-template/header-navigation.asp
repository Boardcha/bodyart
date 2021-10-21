<%
'=== Check if items are stored in wishlist // for navigation
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.Prepared = true
objCmd.CommandText = "SELECT wishlist.ID, wishlist.custID FROM wishlist WHERE wishlist.custID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
Set rsNavWishlist = objCmd.Execute()

'==== Check if items are stored in saved searches // for navigation
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM tbl_customer_searches  WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
Set rsNavSavedSearches = objCmd.Execute()

'==== Check if items are stored in the waiting list // for navigation
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TBLWaitingList.ID FROM TBLWaitingList WHERE customerID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
Set rsNavWaitingList = objCmd.Execute()
%>
</head>
<body>
        <!-- Google Tag Manager (noscript) -->
<noscript><iframe src="https://www.googletagmanager.com/ns.html?id=GTM-W46N98J"
        height="0" width="0" style="display:none;visibility:hidden"></iframe></noscript>
        <!-- End Google Tag Manager (noscript) -->
        <% if var_display_coupon_code <> "" then %>
        <div class="bg-dark text-light rounded-0 p-1 m-0 small text-center border-bottom border-secondary" role="alert">
                <strong>
                        <%= var_display_coupon_amount %>% OFF!</strong> Use code <strong>
                        <%= var_display_coupon_code %></strong> at checkout. Ends
                <%=MonthName(Month(var_display_end_date),1)%>&nbsp;
                <%= Day(var_display_end_date)%>
        </div>
        <% end if %>
               <!-- LOST / HELP Modal -->
          <div class="modal fade" id="newsitehelp" tabindex="-1" role="dialog" aria-labelledby="newsitehelpLabel"
          aria-hidden="true">
          <div class="modal-dialog" role="document">
                  <div class="modal-content">
                          <div class="modal-body">
                               <img class="img-fluid" src="/images/mobile-header.jpg">
                               <div class="small my-1"><i>Example of mobile navigation menu</i></div>
                               Here's some information to help you locate what you need on mobile devices. Everything is pretty much in the same location as before with a few minor tweaks.
                               <i class="fa fa-bars d-block mt-3"></i>
                               The menu icon brings up all the categories to browse all our products
                               <i class="fa fa-search d-block mt-3"></i> The search icon will bring up the search bar and all the advanced filter options
                               <i class="fa fa-user d-block mt-3"></i> The user icon will bring up the sign in / registration window. OR if you're already signed in, it will bring up the account menu.
                          </div>
                          <div class="modal-footer">
                                        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
                                </div>
                  </div>
          </div>
  </div><!-- LOST / HELP modal -->
        <header class="container-fluid header-bar mr-0 pr-0" id="page-top">
                <div class="row w-100">
                        <a class="col-2 col-sm-1 d-lg-none py-3 py-md-4 px-0 bg-dark my-auto text-light text-center" href="#">
                                <i class="fa fa-bars fa-lg hamburger" data-toggle="collapse" data-target="#mobilemenu"
                                        aria-controls="mobilemenu" aria-expanded="false" aria-label="Toggle navigation"></i>
                        </a>
                        <div class="col-5 col-sm-5 col-lg-2 py-3 my-auto">
                                <a href="/">
                                        <img src="/images/baf-logo-solid-white.png" class="img-fluid">
                                </a>
                        </div>
                        <div class="col-5 col-sm-6 col-lg-10 px-0 my-auto text-right">
                                <form class="form-inline d-none d-lg-inline pr-4" action="products.asp" method="get">
                                        <input class="form-control mr-2 bg-lightgrey shadow-none border-secondary text-dark"
                                                name="keywords" id="desktop-keywords" type="search" placeholder="Search">
                                        <button class="btn btn-sm my-2 text-light border-none bg-dark" type="submit"
                                                name="btn-search">
                                                <i class="fa fa-search"></i>
                                                </h3>
                                        </button>
                                        
                                        <button class="shadow-none d-none d-lg-inline btn btn-sm my-2 text-light border-none bg-dark ml-3 header-menu-open" id="toggle-filters-pc" type="button" data-toggle="collapse"
                                                data-target="#filters" aria-controls="filters" aria-expanded="false"
                                                aria-label="Toggle filters">
                                                <i class="fa fa-filter mr-1"></i> Filter Products
                                        </button>
                                </form>
                                <a class="mx-2 text-light py-3 d-lg-none" href="#">
                                        <i class="fa fa-search fa-lg mobile-search-icon" id="toggle-filters-mobile" data-toggle="collapse"
                                                data-target="#filters" aria-controls="filters" aria-expanded="false"
                                                aria-label="Toggle filters"></i>
                                </a>
                                <% If not rsGetUser.EOF and request.cookies("ID") <> "" then %>
                                <a class="dropdown dropdown-toggle text-light d-none d-lg-inline-block pr-3 py-2"  href="#"
                                id="accountDropdown" data-toggle="collapse" data-target="#accountmenu-bar" aria-controls="accountmenu-bar" aria-expanded="false" aria-label="Toggle account navigation">
                                        My Account
                                </a>
                                <a class="dropdown text-light d-lg-none mx-2" href="#" id="mobileaccountDropdown" data-toggle="collapse" data-target="#accountmenu-bar" aria-controls="accountmenu-bar" aria-expanded="false" aria-label="Toggle account navigation">
                                        <i class="fa fa-user fa-lg pr-xxs-1 px-xs-2"></i>
                                </a>
                                <% else ' not logged in %>
                                <a class="mx-2 text-light" data-toggle="modal" data-target="#signin" href="#">
                                        <span class="d-none d-md-inline-block pr-3">Sign In / Register</span>
                                        <i class="fa fa-user fa-lg d-md-none pr-xxs-1 pl-xs-2 pr-xs-3"></i>
                                </a>
                                <% end if %>
                                <div class="dropdown d-inline-block">
                                        <!-- DESKTOP CART BUTTON -->
                                        <a class="btn btn-cart d-none d-md-inline-block shadow-none rounded-0"  href="/cart.asp">
                                                <i class="fa fa-shopping-cart fa-lg pr-1" aria-hidden="true"></i>
                                                <span class="badge badge-pill badge-danger cart-count"></span>
                                        </a><button type="button" class="btn btn-cart d-none d-md-inline-block shadow-none rounded-0 px-2 btn-cart-load" style="margin-left:1px"
                                        data-toggle="dropdown" role="button"                                                 aria-haspopup="true" aria-expanded="false">
                                                <i class="fa fa-chevron-down"></i>
                                         </button>
                                        <!-- MOBILE CART BUTTON -->
                                        <a class="text-light d-block d-md-none pl-xxs-1 btn-cart-load" href="#" data-toggle="dropdown"
                                                aria-haspopup="true" aria-expanded="false">
                                                <i class="fa fa-shopping-cart fa-lg pr-1"></i>
                                                <span class="badge badge-pill badge-danger cart-count"></span>
                                        </a>
                                        <div class="dropdown-menu dropdown-menu-right bg-light p-3 shadow-lg border-bottom text-dark cart-preview"
                                                aria-labelledby="cartDropdown">
                                                <div class="cart-mini-load">
                                                        <i class="fa fa-spinner fa-2x fa-spin text-secondary"></i>
                                                </div>
                                                <div class="row">
                                                        <div class="col-lg-12 col-sm-12 col-12 text-center pt-2">
                                                                <a class="btn btn-purple btn-block" href="/cart.asp">View Cart / Checkout</a>
                                                        </div>
                                                </div>
                                        </div>
                                </div>
                        </div>
                </div>
                </div>
        </header>
        <% If not rsGetUser.EOF and request.cookies("ID") <> "" then %>
        <div class="collapse bg-dark border-bottom border-top border-secondary small text-sm-left text-lg-right pl-2" id="accountmenu-bar">
                <a class="p-2 d-block d-lg-inline-block text-light border-left border-right border-secondary account-nav-links" id="account-order-history" href="/account.asp">Order history</a>
                <a class="p-2 d-block d-lg-inline-block text-light border-right border-secondary account-nav-links" id="account-credit-cards" href="/account-billing.asp">Credit cards</a>
                <a class="p-2 d-block d-lg-inline-block text-light border-right border-secondary account-nav-links" id="account-shipping-addresses"  href="/account-shipping.asp">Shipping addresses</a>
                <% if not rsNavWishlist.eof then %>
                <a class="p-2 d-block d-lg-inline-block text-light border-right border-secondary account-nav-links" id="account-wishlist"  href="/wishlist.asp">Wishlist</a>
                <% end if %>
                <a class="p-2 d-block d-lg-inline-block text-light border-right border-secondary account-nav-links" id="account-credits-rewards"  href="/account-credits.asp">Account credits &amp; rewards</a>
                <% if not rsNavSavedSearches.eof then %>
                <a class="p-2 d-block d-lg-inline-block text-light border-right border-secondary account-nav-links" id="account-saved-searches"  href="/account-searches.asp">Saved searches</a>
                <% end if %>
                <% if not rsNavWaitingList.eof then %>
                <a class="p-2 d-block d-lg-inline-block text-light border-right border-secondary account-nav-links" id="account-waiting-list"  href="/account-waiting-list.asp">Waiting List</a>
                <% end if %>
                <a class="p-2 d-block d-lg-inline-block text-light border-right border-secondary account-nav-links" id="account-manage-profile"  href="/account-profile.asp">Profile</a>
                <a class="p-2 d-block d-lg-inline-block mr-2 text-light account-nav-links" id="account-logout"  href="/sign-out.asp">Logout</a>
        </div>
        <% end if ' not logged in %>
        <nav class="navbar navbar-expand-lg p-0 top-navbar">
                <div class="navbar-nav w-100 px-">
                        <div class="collapse navbar-collapse" id="mobilemenu">
                                <ul class="navbar-nav mr-auto" style="background-color: #2E2E2E">
                                        <li class="nav-item p-1">
                                                <a class="nav-link text-light py-2 pl-3 header-menu-link" href="/products.asp?new=Yes" id="new-items">NEW<span
                                                                class="d-md-none"> Jewelry</span>
                                                </a>
                                        </li>
                                        <li class="nav-item dropdown position-static  p-1">
                                                <a class="nav-link dropdown-toggle text-light py-2 px-3 px-lg-1 header-menu-open" href="#" id="saleDropdown"
                                                        role="button" data-toggle="dropdown" aria-haspopup="true"
                                                        aria-expanded="false">
                                                        SALES
                                                </a>
                                                <div class="dropdown-menu w-100 m-0 border-0 rounded-0 nav-dropdown"
                                                        aria-labelledby="saleDropdown">
                                                        <div class="container-fluid pb-3">
                                                                <div class="row">

                                                                        <div class="col  mx-lg-auto">
                                                                               
                                                                                <a class="ropdown-item track-sales-dropdown" href="/products.asp?discount=all" id="sales-all">
                                                                                        View all sale items
                                                                                </a>
                                                                                <a class="dropdown-item track-sales-dropdown" href="/products.asp?discount=5-20" id="sales-5to20-discount">
                                                                                        5% to 20% off
                                                                                </a>
                                                                                <a class="dropdown-item track-sales-dropdown" href="/products.asp?discount=25-45" id="sales-25to45-discount">
                                                                                        25% to 45% off
                                                                                </a>
                                                                                <a class="dropdown-item track-sales-dropdown" href="/products.asp?discount=50-70" id="sales-50to70-discount">
                                                                                        50% to 70% off
                                                                                </a>
                                                                                <a class="dropdown-item track-sales-dropdown" href="/products.asp?discount=75-90" id="sales-75up-discount">
                                                                                        75% + off
                                                                                </a>        
                                                                        </div>
                                                                                                               </div>
                                                        </div>
                                                </div>
                                        </li>
                                        <li class="nav-item dropdown position-static p-1">
                                                <a class="nav-link dropdown-toggle text-light py-2  px-3 px-lg-1 header-menu-open" href="#" id="basicsDropdown"
                                                        role="button" data-toggle="dropdown" aria-haspopup="true"
                                                        aria-expanded="false">
                                                        Basics
                                                </a> 
                                                <div class="small text-secondary px-3 d-lg-none">Balls, rings, barbells, labrets, circulars &amp; curves</div>
                                                <div class="dropdown-menu w-100 m-0 border-0 rounded-0 nav-dropdown"
                                                        aria-labelledby="basicsDropdown">
                                                        <div class="container-fluid pb-3">
                                                                <div class="row">
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto">
                                                                                <a class="h6 d-block pb-1 border-bottom border-secondary track-basics-dropdown track-nav-loose-ends" href="/products.asp?jewelry=balls" id="image-loose-ends">
                                                                                    <img data-src="/images/navigation/navi-400-ends-o-rings.jpg" class="img-fluid img-75 d-block lazyload" />
                                                                                        Balls, ends, &amp; beads <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                all</span></span>
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-loose-ends" href="/products.asp?jewelry=basics&amp;jewelry=balls" id="ends-basic">
                                                                                                Basic ends
                                                                                        </a>
                                                                                        <a class="dropdown-item track-basics-dropdown track-nav-loose-ends" href="/products.asp?threading=Threadless&amp;jewelry=balls" id="ends-threadless">
                                                                                                Threadless ends
                                                                                        </a>
                                                                                        <a class="dropdown-item track-basics-dropdown track-nav-loose-ends" href="/products.asp?jewelry=balls&amp;threading=Internally+threaded" id="ends-internal">
                                                                                                Internally threaded ends
                                                                                        </a>
                                                                                        <a class="dropdown-item track-basics-dropdown track-nav-loose-ends" href="/products.asp?jewelry=balls&amp;threading=Externally+threaded" id="ends-external">
                                                                                                Externally threaded ends
                                                                                        </a>
                                                                                        <a class="dropdown-item track-basics-dropdown track-nav-loose-ends" href="/products.asp?jewelry=balls&material=solid+rose+gold&material=solid+white+gold&material=solid+yellow+gold" id="ends-gold">
                                                                                                Gold ends
                                                                                        </a>
                                                                                        <a class="dropdown-item track-basics-dropdown track-nav-loose-ends" href="/products.asp?jewelry=beads" id="ends-beads">
                                                                                                Replacement beads
                                                                                        </a>
                                                                                        <a class="dropdown-item track-basics-dropdown track-nav-loose-ends" href="/products.asp?jewelry=charms" id="ends-charms">
                                                                                                Charms
                                                                                        </a>
                                                                                        <a class="dropdown-item track-basics-dropdown track-nav-loose-ends" href="/products.asp?jewelry=orings" id="ends-orings">
                                                                                                O-rings
                                                                                        </a>
                                                                        </div>
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto">
                                                                                <a class="h6 d-block pb-1 border-bottom border-secondary  track-basics-dropdown track-nav-captives" href="/products.asp?jewelry=captive" id="image-captives">
                                                                                     <img data-src="/images/navigation/navi-400-rings.jpg" class="img-fluid img-75 d-block lazyload" />
                                                                                        Rings <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                all</span></span>
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-captives" href="/products.asp?jewelry=basics&amp;jewelry=captive" id="captives-basic">
                                                                                        Basic rings
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-captives" href="/products.asp?jewelry=clicker" id="captives-clickers">
                                                                                        Clickers
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-captives" href="/products.asp?keywords=seamless&amp;jewelry=captive" id="captives-seamless">
                                                                                        Seamless rings
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-captives" href="/products.asp?jewelry=captive-cbr" id="captives-cbrs">
                                                                                        Captive Bead Rings
                                                                                </a>
                                                                        </div>
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto">
                                                                                <a class="h6 d-block pb-1 border-bottom border-secondary  track-basics-dropdown track-nav-labret" href="/products.asp?jewelry=labret" id="image-labret">
                                                                                        <img data-src="/images/navigation/navi-400-lip-labret.jpg" class="img-fluid img-75 d-block lazyload" />
                                                                                        Labret / Lip Jewelry <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                all</span></span>
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-labret" href="/products.asp?jewelry=basics&amp;jewelry=labret" id="labrets-basic">
                                                                                        Basic labrets
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-labret" href="/products.asp?jewelry=labret-design" id="labrets-design-ends">
                                                                                        Design end labrets
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-labret" href="/products.asp?jewelry=labret&threading=Threadless" id="labrets-threadless">
                                                                                        Threadless labrets
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-labret" href="/products.asp?jewelry=captive&amp;piercing=labret" id="labrets-rings">
                                                                                        Lip rings
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-labret" href="/products.asp?jewelry=labret-stretched" id="labret-stretched">
                                                                                        Stretched labret
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-labret" href="/products.asp?keywords=retainer&amp;piercing=labret" id="labrets-retainers">
                                                                                        Lip retainers
                                                                                </a>
                                                                        </div>
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto">
                                                                                <a class="h6 d-block pb-1 border-bottom border-secondary  track-basics-dropdown track-nav-barbells"  href="/products.asp?jewelry=barbell" id="image-barbells">
                                                                                    <img data-src="/images/navigation/navi-400-straight-barbells.jpg" class="img-fluid img-75 d-block lazyload" />
                                                                                        Straight Barbells <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                all</span></span>
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-barbells" href="/products.asp?jewelry=basics&amp;jewelry=barbell" id="barbells-basic">
                                                                                        Basic barbells
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-barbells" href="/products.asp?keywords=helix&amp;jewelry=barbell" id="barbells-cartilage">
                                                                                        Cartilage barbells
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-barbells" href="/products.asp?keywords=industrial&amp;jewelry=barbell" id="barbells-industrials">
                                                                                        Industrial barbells
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-barbells" href="/products.asp?jewelry=barbell&amp;piercing=nipple" id="barbells-nipple">
                                                                                        Nipple barbells
                                                                                </a>
                                                                                <a class="dropdown-item track-basics-dropdown track-nav-barbells" href="/products.asp?keywords=tongue&amp;jewelry=barbell" id="barbells-tongue">
                                                                                        Tongue barbells
                                                                                </a>
                                                                        </div>
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto">
                                                                                        
                                                                                        <a class="h6 track-basics-dropdown track-nav-misc" href="/products.asp?jewelry=circular&jewelry=curved&jewelry=twists" id="image-circulars-curves-twists">
                                                                                 
                                                                                        <img data-src="/images/navigation/navi-400-curved-circular-barbells.jpg" class="img-fluid img-75 d-block lazyload" /></a>
                                                                                <a class="h6 d-block pb-1 border-bottom border-secondary track-basics-dropdown track-nav-misc" href="/products.asp?jewelry=circular" id="circulars">
                                                                                        Circular Barbells <span class="badge badge-secondary   ml-xl-4 font-weight-normal">View
                                                                                                all</span>
                                                                                </a>
                                                                                <a class="h6 d-block pb-1 border-bottom border-secondary track-basics-dropdown track-nav-misc" href="/products.asp?jewelry=curved" id="curves-view-all">
                                                                                        Curved Barbells <span class="badge badge-secondary   ml-xl-4 font-weight-normal">View
                                                                                                all</span>
                                                                                </a>
                                                                                <a class="h6 d-block pb-1 border-bottom border-secondary track-basics-dropdown track-nav-misc" href="/products.asp?jewelry=twists" id="twists">
                                                                                        Twists <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                all</span></span>
                                                                                </a>
                                                                        </div>
                                                                </div>
                                                        </div>
                                                </div>
                                         </li>
                                         <li class="nav-item dropdown position-static  p-1">
                                                <a class="nav-link dropdown-toggle text-light py-2  px-3 px-lg-1 header-menu-open" href="#" id="septumDropdown"
                                                        role="button" data-toggle="dropdown" aria-haspopup="true"
                                                        aria-expanded="false">
                                                        Septum/Nose</a>
                                                        <div class="small text-secondary px-3 d-lg-none">Clickers + all septum jewelry, nose rings &amp; studs + more</div>
                                                <div class="dropdown-menu w-100 m-0 border-0 rounded-0 nav-dropdown"
                                                        aria-labelledby="septumDropdown">
                                                        <div class="container-fluid pb-3">
                                                                <div class="row">
                                                                        <div class="col-6 order-1 order-lg-1 col-lg-3">
                                                                                        <a class="h6 d-block pb-lg-1 track-septum-dropdown track-septums"
                                                                                        href="/products.asp?jewelry=septum" id="image-septum">
                                                                                        <img  data-src="/images/navigation/navi-400-septum.jpg" class="img-fluid img-75 lazyload" /></a>
                                                                        </div>
                                                                        <div class="col-6 order-3 order-lg-2 col-lg-3">
                                                                                <a class="h6 d-block pb-1 border-bottom border-secondary track-septum-dropdown track-septums"
                                                                                        href="/products.asp?jewelry=septum" id="septum-view-all-link">
                                                                                        Septum Jewelry <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                all</span></span>
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-septums" href="/products.asp?keywords=clicker&jewelry=septum" id="septum-clickers">
                                                                                        Septum clickers
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-septums" href="/products.asp?jewelry=basics&amp;jewelry=septum" id="basic-septums">
                                                                                        Basic septums
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-septums" href="/products.asp?keywords=seamless&piercing=septum" id="septum-seamless-rings">
                                                                                        Seamless rings
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-septums" href="/products.asp?jewelry=septum-captive" id="septum-captives">
                                                                                        Septum captive rings
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-septums" href="/products.asp?material=gold&jewelry=septum" id="gold-septums">
                                                                                        Gold septums
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-septums" href="/products.asp?keywords=pincher&amp;piercing=septum" id="septum-pinchers">
                                                                                        Pinchers
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-septums" href="/products.asp?jewelry=circular&amp;piercing=septum" id="septum-circulars">
                                                                                        Circular barbells
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-septums" href="/products.asp?jewelry=septum-spike" id="septum-spikes">
                                                                                        Septum tusks &amp; spikes
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-septums" href="/products.asp?keywords=retainer&piercing=septum" id="septum-retainers">
                                                                                        Septum hiders / retainers
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-septums" href="/products.asp?keywords=plug&piercing=septum" id="septum-plugs">
                                                                                        Septum plugs
                                                                                </a>
                                                                        </div>
                                                                        <div class="col-6 order-2 order-lg-3 col-lg-3">
                                                                                        <a class="h6 d-block pb-lg-1" href="/products.asp?jewelry=nose-ring" id="image-nose-jewelry">
                                                                                        <img data-src="/images/navigation/navi-400-nose.jpg" class="img-fluid img-75 lazyload" />
                                                                                </a>
                                                                        </div>
                                                                        <div class="col-6 order-4 order-lg-4 col-lg-3">
                                                                                <a class="h6 d-block pb-1 border-bottom border-secondary track-septum-dropdown track-nose" href="/products.asp?jewelry=nose-ring" id="nostril-view-all-link">
                                                                                        Nose rings & studs <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                all</span></span>
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-nose" href="/products.asp?jewelry=nose-hoop" id="nose-hoops">
                                                                                        Nose rings & hoops
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-nose" href="/products.asp?keywords=nosescrew&jewelry=nose-ring" id="nosescrews">
                                                                                        Nose screws &amp; L bends
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-nose" href="/products.asp?keywords=nosebones&jewelry=nose-ring" id="nosebones">
                                                                                        Nose studs
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-nose" href="/products.asp?jewelry=chains-short" id="chains-short">
                                                                                        Short nose chains
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-nose" href="/products.asp?jewelry=basics&amp;jewelry=nose-ring" id="basic-nostril">
                                                                                        Basic nose jewelry
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-nose" href="/products.asp?jewelry=nose-ring&amp;material=gold" id="gold-nostril">
                                                                                        Gold nose jewelry
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-nose" href="/products.asp?jewelry=nose-threadless" id="nostril-threadless">
                                                                                        Threadless nose studs with backs
                                                                                </a>
                                                                                <a class="dropdown-item track-septum-dropdown track-nose" href="/products.asp?keywords=retainer&jewelry=nose-ring" id="nose-retainers">
                                                                                        Clear & skin tone nose hiders
                                                                                </a>
                                                                        </div>
                                                                </div>
                                                        </div>
                                                </div>
                                        </li>
                                        <li class="nav-item dropdown position-static p-1">
                                                <a class="nav-link text-light py-2 px-3 px-lg-1 dropdown-toggle header-menu-open" href="#" id="plugsDropdown"
                                                        role="button" data-toggle="dropdown" aria-haspopup="true"
                                                        aria-expanded="false">
                                                        Plugs <span class="d-lg-none">&amp; Tunnels / Stretching</span><span class="d-none d-lg-inline"> / Stretching</span></a>
                                                        <div class="small text-secondary px-3 d-lg-none">Jewelry for stretched ears (up to 3") &amp; stretching tapers</div>
                                                <div class="dropdown-menu w-100 m-0 border-0 rounded-0 nav-dropdown"
                                                        aria-labelledby="plugsDropdown">
                                                        <div class="container-fluid pb-3">
                                                                <div class="row">
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto">
                                                                                        <a class="h6 d-block pb-1 track-plugs-dropdown" href="/products.asp?jewelry=plugs" id="image-plugs">  
                                                                                            <img  data-src="/images/navigation/navi-400-plugs.jpg" class="img-fluid img-75 lazyload" /> 
                                                                                            <div class="d-block">Plugs &amp; Tunnels <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                        all</span></span>
                                                                                                </div>
                                                                                        </a>
                                                                                        <a class="h6 d-block pb-1 track-plugs-dropdown" href="/products.asp?jewelry=saddle" id="image-saddles">  
                                                                                                <img  data-src="/images/navigation/navi-400-saddles.jpg" class="img-fluid img-75 lazyload" /> 
                                                                                                <div class="d-block">Saddles<span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                            all</span></span>
                                                                                                    </div>
                                                                                            </a>
                                                                                            <a class="h6 d-block pb-1 track-plugs-dropdown" href="/products.asp?jewelry=tapers" id="image-tapers"><img data-src="/images/navigation/navi-400-stretching.jpg"
                                                                                                class="img-fluid 
                                                                                                img-75 d-block lazyload">
                                                                                                <div class="d-block">
                                                                                                Stretching Tools <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                        all</span></span>
                                                                                                </div>
                                                                                        </a>
                                                                        </div>
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto">
                                                                                <h6 class="d-lg-block pb-1 border-bottom border-secondary">18g
                                                                                        through 00g</h6>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=18g" id="18g-plugs">18g
                                                                                        &nbsp;&nbsp;(1mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=16g" id="16g-plugs">16g
                                                                                        &nbsp;&nbsp;(1.2mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=14g" id="14g-plugs">14g
                                                                                        &nbsp;&nbsp;(1.6mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=12g" id="12g-plugs">12g
                                                                                        &nbsp;&nbsp;(2mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=10g" id="10g-plugs">10g
                                                                                        &nbsp;&nbsp;(2.5mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=8g" id="8g-plugs">8g
                                                                                        &nbsp;&nbsp;(3mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=7g" id="7g-plugs">7g
                                                                                        &nbsp;&nbsp;(3.5mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=6g" id="6g-plugs">6g
                                                                                        &nbsp;&nbsp;(4mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=5g" id="5g-plugs">5g
                                                                                        &nbsp;&nbsp;(4.5mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=4g" id="4g-plugs">4g
                                                                                        &nbsp;&nbsp;(5mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=3g" id="3g-plugs">3g
                                                                                                &nbsp;&nbsp;(5.5mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2g" id="2g-plugs">2g
                                                                                        &nbsp;&nbsp;(6mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1g" id="1g-plugs">1g
                                                                                        &nbsp;&nbsp;(7mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=0g" id="0g-plugs">0g
                                                                                        &nbsp;&nbsp;(8mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=00g&amp;gauge=00g%2F9mm&amp;gauge=00g%2F9.5mm&amp;gauge=00g%2F10mm" id="All-00g-plugs">00g
                                                                                        &nbsp;&nbsp;(9mm - 10mm)</a>
                                                                        </div>
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto mt-3 mt-sm-0">
                                                                                <h6 class="d-lg-block pb-1 border-bottom border-secondary">7/16"
                                                                                        through 1-5/16"</h6>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=7%2F16%22" id="7/16-plugs">7/16&quot;
                                                                                        &nbsp;&nbsp;(11mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1%2F2%22" id="1/2-plugs">1/2&quot;
                                                                                        &nbsp;&nbsp;(12.5mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=9%2F16%22" id="9/16-plugs">9/16&quot;
                                                                                        &nbsp;&nbsp;(14mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=5%2F8%22" id="5/8-plugs">5/8&quot;
                                                                                        &nbsp;&nbsp;(16mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=11%2F16%22" id="11/16-plugs">11/16&quot;
                                                                                        &nbsp;&nbsp;(18mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=3%2F4%22" id="3/4-plugs">3/4&quot;
                                                                                        &nbsp;&nbsp;(19mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=13%2F16%22" id="13/16-plugs">13/16&quot;
                                                                                        &nbsp;&nbsp;(20mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=7%2F8%22" id="7/8-plugs">7/8&quot;
                                                                                        &nbsp;&nbsp;(22mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=15%2F16%22" id="15/16-plugs">15/16&quot;
                                                                                        &nbsp;&nbsp;(24mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1%22"
                                                                                        class="sf-with-ul" id="1-inch-plugs">1&quot;
                                                                                        &nbsp;&nbsp;(25mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-1%2F16%22" id="1-1/6-plugs">1-1/16&quot;
                                                                                        &nbsp;&nbsp;(27mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-1%2F8%22" id="1-1/8-plugs">1-1/8&quot;
                                                                                        &nbsp;&nbsp;(28.5mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-1%2F4%22" id="1-1/4-plugs">1-1/4&quot;
                                                                                        &nbsp;&nbsp;(32mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-5%2F16%22" id="1-5/16-plugs">1-5/16&quot;
                                                                                        &nbsp;&nbsp;(33mm)</a>
                                                                        </div>
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto mt-3 mt-md-0">
                                                                                <h6 class="d-lg-block pb-1 border-bottom border-secondary">1-3/8"
                                                                                        through 3"</h6>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-3%2F8%22" id="1-3/8-plugs">1-3/8&quot;
                                                                                        &nbsp;&nbsp;(35mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-1%2F2%22" id="1-1/2-plugs">1-1/2&quot;
                                                                                        &nbsp;&nbsp;(38mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-5%2F8%22" id="1-5/8-plugs">1-5/8&quot;
                                                                                        &nbsp;&nbsp;(41mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-3%2F4%22" id="1-3/4-plugs">1-3/4&quot;
                                                                                        &nbsp;&nbsp;(44mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-7%2F8%22" id="1-7/8-plugs">1-7/8&quot;
                                                                                        &nbsp;&nbsp;(48mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2%22" id="2-inch-plugs">2&quot;
                                                                                        &nbsp;&nbsp;(51mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2-1%2F8%22" id="2-1/8-plugs">2-1/8&quot;
                                                                                        &nbsp;&nbsp;(54mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2-1%2F4%22" id="2-1/4-plugs">2-1/4&quot;
                                                                                        &nbsp;&nbsp;(57mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2-1%2F2%22" id="2-1/2-plugs">2-1/2&quot;
                                                                                        &nbsp;&nbsp;(63.5mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2-5%2F8%22" id="2-5/8-plugs">2-5/8&quot;
                                                                                        &nbsp;&nbsp;(67mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2-3%2F4%22" id="2-3/4-plugs">2-3/4&quot;
                                                                                        &nbsp;&nbsp;(70mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2-7%2F8%22" id="2-7/8-plugs">2-7/8&quot;
                                                                                        &nbsp;&nbsp;(73mm)</a>
                                                                                <a class="dropdown-item track-plugs-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=3%22" id="3-inch-plugs">3&quot;
                                                                                        &nbsp;&nbsp;(76mm)</a>
                                                                        </div>
                                                                         <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto mt-3 mt-lg-0">
                                                                </div>
                                                        </div>
                                                </div>
                                        </li>
                                        <li class="nav-item dropdown position-static p-1">
                                                <a class="nav-link dropdown-toggle text-light py-2  px-3 px-lg-1 header-menu-open" href="#" id="OtherJewelryDropdown"
                                                        role="button" data-toggle="dropdown" aria-haspopup="true"
                                                        aria-expanded="false">
                                                        Other Jewelry</a>
                                                        <div class="small text-secondary px-3 d-lg-none">Hanging designs, weights, navel, nipple, earrings, &amp; necklaces</div>
                                                <div class="dropdown-menu w-100 m-0 border-0 rounded-0 nav-dropdown"
                                                        aria-labelledby="OtherJewelryDropdown">
                                                        <div class="container-fluid pb-3">
                                                                <div class="row">
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto">
                                                                                <a class="h6 d-block pb-1 border-bottom border-secondary track-other-jewelry-dropdown track-earrings"  href="/products.asp?jewelry=earring" id="image-earrings"><img data-src="/images/navigation/navi-400-earrings.jpg"
                                                                                class="img-fluid 
                                                                                               img-75 d-block lazyload">
                                                                                
                                                                                Earrings<span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                        all</span></span>
                                                                        </a>
                                                                        
                                                                        <a class="dropdown-item track-other-jewelry-dropdown track-earrings" href="/products.asp?jewelry=earring-stud" id="earring-stud">
                                                                                Earring studs
                                                                        </a>
                                                                        <a class="dropdown-item track-other-jewelry-dropdown track-earrings" href="/products.asp?jewelry=earring-dangle" id="earring-dangle">
                                                                                Earring hoops & hanging
                                                                        </a>
                                                                        
                                                                        <a class="dropdown-item track-other-jewelry-dropdown track-earrings" href="/products.asp?jewelry=earring-huggies" id="earring-huggie">
                                                                                Earring huggies
                                                                        </a>
                                                                


                                                                        
                                                                </div>
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto">
                                                                                        <a class="h6 d-block pb-1 border-bottom border-secondary track-other-jewelry-dropdown track-hanging"  href="/products.asp?jewelry=Hanging+Designs" id="image-hanging"><img data-src="/images/navigation/navi-400-hanging-designs.jpg"
                                                                                        class="img-fluid 
                                                                                                       img-75 d-block lazyload">
                                                                                        
                                                                                        Hanging Jewelry <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                all</span></span>
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-hanging" href="/products.asp?jewelry=ornate" id="hanging-ornate">
                                                                                        Ornate
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-hanging" href="/products.asp?jewelry=hoop" id="hanging-hoops">
                                                                                        Hoops
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-hanging" href="/products.asp?jewelry=spiral&jewelry=coils" id="hanging-spiral">
                                                                                        Spirals &amp; Coils
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-hanging" href="/products.asp?jewelry=plugloops" id="hanging-plug-hoops">
                                                                                        Plug Hoops
                                                                                </a>
                                                                                <a class="h6 d-block mt-2 pb-1 border-bottom border-secondary track-other-jewelry-dropdown track-hanging" href="/products.asp?jewelry=weight" id="image-weights">
                                                                                Weights (View all) <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                        all</span></span>
                                                                        </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-hanging" href="/products.asp?jewelry=weight-light" id="weights-light">
                                                                                        Light (5g - 15g)
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-hanging" href="/products.asp?jewelry=weight-medium" id="weights-medium">
                                                                                        Medium (16g - 25g)
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-hanging" href="/products.asp?jewelry=weight-heavy" id="weights-heavy">
                                                                                        Heavy (26g - 40g)
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-hanging" href="/products.asp?jewelry=weight-super-heavy" id="weights-super-heavy">
                                                                                        Extra heavy (40g +)
                                                                                </a>          


                                                                                
                                                                        </div>
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto">
                                                                                        
                                                                                        <a class="h6 d-block pb-1 border-bottom border-secondary track-other-jewelry-dropdown track-navel" href="/products.asp?jewelry=belly" id="image-navels"><img data-src="/images/navigation/navi-400-navel.jpg"
                                                                                        class="img-fluid 
                                                                                                img-75 d-block lazyload">
                                                                                        Belly Button Jewelry <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                all</span></span>
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-navel" href="/products.asp?jewelry=belly-simple" id="navels-simple">
                                                                                        Belly jewelry (no dangles)
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-navel" href="/products.asp?jewelry=belly-dangle" id="navels-dangle">
                                                                                        Belly jewelry (with dangles)
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-navel" href="/products.asp?keywords=retainer&jewelry=belly" id="navel-retainers">
                                                                                        Retainers
                                                                                </a>
                                                                        </div>
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto">
                                                                                        <a class="h6 d-block pb-1 border-bottom border-secondary track-other-jewelry-dropdown track-nipple" href="/products.asp?jewelry=nipple" id="image-nipple-jewelry"><img data-src="/images/navigation/navi-400-nipple-jewelry.jpg"
                                                                                        class="img-fluid 
                                                                                                       img-75 d-block lazyload">
                                                                                        Nipple Jewelry <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                all</span></span>
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-nipple" href="/products.asp?jewelry=barbell&piercing=nipple" id="nipple-barbells">
                                                                                        Nipple barbells
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-nipple" href="/products.asp?keywords=clicker&jewelry=nipple" id="nipple-clickers">
                                                                                        Hinged nipple jewelry
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-nipple" href="/products.asp?jewelry=nipple-capcir" id="nipple-captives-and-circulars">
                                                                                        Captives &amp; Circulars
                                                                                </a> 
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-nipple" href="/products.asp?jewelry=nipple-shield&amp;jewelry=nipple-stirrup" id="nipple-stirrup">
                                                                                        Shields &amp; stirrups
                                                                                </a>
                                                                              
                                                                                                                                                     
                                                                                    
                                                                        </div>
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2 mx-lg-auto">
                                                                                        <a class="h6 d-block pb-1 border-bottom border-secondary track-other-jewelry-dropdown track-regular-jewelry" href="/products.asp?jewelry=bracelet&amp;jewelry=earring&amp;jewelry=necklace&amp;jewelry=finger-ring" id="image-all-regular-jewelry"><img data-src="/images/navigation/navi-400-regular-jewelry.jpg"
                                                                                        class="img-fluid 
                                                                                                img-75 d-block lazyload">
                                                                                        All regular jewelry <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                all</span></span>
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-regular-jewelry" href="/products.asp?jewelry=finger-ring" id="finger-rings">
                                                                                        Finger rings
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-regular-jewelry" href="/products.asp?jewelry=necklace" id="necklaces">
                                                                                        Necklaces &amp; Pendants
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-regular-jewelry" href="/products.asp?jewelry=chains-necklace" id="chains-necklace">
                                                                                        Necklace chains
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-regular-jewelry" href="/products.asp?jewelry=bracelet" id="bracelets">
                                                                                        Bracelets
                                                                                </a>
                                                                                <a class="dropdown-item track-other-jewelry-dropdown track-regular-jewelry" href="/products.asp?keywords=ear+cuff" id="ear-cuffs">
                                                                                        Ear cuffs
                                                                                </a>
                                                                        </div>
                                                                </div>
                                                        </div>
                                                </div>
                                        </li>
                                        <li class="nav-item dropdown position-static p-1">
                                                <a class="nav-link dropdown-toggle text-light py-2  px-3 px-lg-1 header-menu-open" href="#" id="gaugesDropdown"
                                                        role="button" data-toggle="dropdown" aria-haspopup="true"
                                                        aria-expanded="false">
                                                        Gauges</a>
                                                        <div class="small text-secondary px-3 d-lg-none">Browse by gauge from 18g to 3" (1mm to 76mm)</div>
                                                <div class="dropdown-menu w-100 m-0 border-0 rounded-0 nav-dropdown"
                                                        aria-labelledby="gaugesDropdown">
                                                        <div class="container-fluid pb-3">
                                                                <div class="row">
                                                                        <div class="col-6 col-sm-3">
                                                                                <h6 class="d-lg-block pb-1 border-bottom border-secondary">18g
                                                                                                        through 00g</h6>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=18g" id="gauge-18g">18g
                                                                                        &nbsp;&nbsp;(1mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=16g" id="gauge-16g">16g
                                                                                        &nbsp;&nbsp;(1.2mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=14g" id="gauge-14g">14g
                                                                                        &nbsp;&nbsp;(1.6mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=12g" id="gauge-12g">12g
                                                                                        &nbsp;&nbsp;(2mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=10g" id="gauge-10g">10g
                                                                                        &nbsp;&nbsp;(2.5mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=8g" id="gauge-8g">8g
                                                                                        &nbsp;&nbsp;(3mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=7g" id="gauge-7g">7g
                                                                                        &nbsp;&nbsp;(3.5mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=6g" id="gauge-6g">6g
                                                                                        &nbsp;&nbsp;(4mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=5g" id="gauge-5g">5g
                                                                                        &nbsp;&nbsp;(4.5mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=4g" id="gauge-4g">4g
                                                                                        &nbsp;&nbsp;(5mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=3g" id="gauge-3g">3g
                                                                                                &nbsp;&nbsp;(5.5mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=2g" id="gauge-2g">2g
                                                                                        &nbsp;&nbsp;(6mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=1g" id="gauge-1g">1g
                                                                                        &nbsp;&nbsp;(7mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=0g" id="gauge-0g">0g
                                                                                        &nbsp;&nbsp;(8mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=00g&amp;gauge=00g%2F9mm&amp;gauge=00g%2F9.5mm&amp;gauge=00g%2F10mm" id="gauge-00g">00g
                                                                                        &nbsp;&nbsp;(9mm - 10mm)</a>
                                                                        </div>
                                                                        <div class="col-6 col-sm-3">
                                                                        <h6 class="d-lg-block pb-1 border-bottom border-secondary">7/16&quot; (11mm)
                                                                                                        through 1&quot; (25mm)</h6>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=7%2F16%22" id="gauge-7/16">7/16&quot;
                                                                                        &nbsp;&nbsp;(11mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=1%2F2%22" id="gauge-1/2">1/2&quot;
                                                                                        &nbsp;&nbsp;(12.5mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=9%2F16%22" id="gauge-9/16">9/16&quot;
                                                                                        &nbsp;&nbsp;(14mm)</a>
                                                                                        <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=15mm" id="gauge-15mm">15mm
                                                                                                &nbsp;&nbsp;</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=5%2F8%22" id="gauge-5/8">5/8&quot;
                                                                                        &nbsp;&nbsp;(16mm)</a>
                                                                                        <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=17mm" id="gauge-17mm">17mm
                                                                                                &nbsp;&nbsp;</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=11%2F16%22" id="gauge-11/16">11/16&quot;
                                                                                        &nbsp;&nbsp;(18mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" class="dropdown-item track-gauges-dropdown"
                                                                                        href="/products.asp?gauge=3%2F4%22" id="gauge-3/4">3/4&quot;
                                                                                        &nbsp;&nbsp;(19mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=13%2F16%22" id="gauge-13/16">13/16&quot;
                                                                                        &nbsp;&nbsp;(20mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=7%2F8%22" id="gauge-7/8">7/8&quot;
                                                                                        &nbsp;&nbsp;(22mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=15%2F16%22" id="gauge-15/16">15/16&quot;
                                                                                        &nbsp;&nbsp;(24mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?gauge=1%22"
                                                                                        class="sf-with-ul" id="gauge-1inch">1&quot;
                                                                                        &nbsp;&nbsp;(25mm)</a>
                                                                        </div>
                                                                        <div class="col-6 col-sm-3 mt-xxs-3 mt-sm-0">
                                                                                <h6 class="d-lg-block pb-1 border-bottom border-secondary">1-1/16&quot; (27mm)
                                                                                        through 1-7/8&quot; (48mm)</h6>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-1%2F16%22" id="1-1/6-plugs">1-1/16&quot;
                                                                                        &nbsp;&nbsp;(27mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-1%2F8%22" id="1-1/8-plugs">1-1/8&quot;
                                                                                        &nbsp;&nbsp;(28.5mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-1%2F4%22" id="1-1/4-plugs">1-1/4&quot;
                                                                                        &nbsp;&nbsp;(32mm)</a>
                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-5%2F16%22" id="1-5/16-plugs">1-5/16&quot;
                                                                                        &nbsp;&nbsp;(33mm)</a> <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-3%2F8%22"
                                                                                        id="1-3/8-plugs">1-3/8&quot;
                                                                                        &nbsp;&nbsp;(35mm)</a>
                                                                                        <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-1%2F2%22" id="1-1/2-plugs">1-1/2&quot;
                                                                                                &nbsp;&nbsp;(38mm)</a>
                                                                                        <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-5%2F8%22" id="1-5/8-plugs">1-5/8&quot;
                                                                                                &nbsp;&nbsp;(41mm)</a>
                                                                                        <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-3%2F4%22" id="1-3/4-plugs">1-3/4&quot;
                                                                                                &nbsp;&nbsp;(44mm)</a>
                                                                                        <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=1-7%2F8%22" id="1-7/8-plugs">1-7/8&quot;
                                                                                                &nbsp;&nbsp;(48mm)</a>
                                                                                        </div>
                                                                                        <div class="col-6 col-sm-3 mt-xxs-3 mt-sm-0">
                                                                                                <h6 class="d-lg-block pb-1 border-bottom border-secondary">2&quot; (51mm)
                                                                                                        through 3&quot; (76mm)</h6>
                                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2%22" id="2-inch-plugs">2&quot;
                                                                                                        &nbsp;&nbsp;(51mm)</a>
                                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2-1%2F8%22" id="2-1/8-plugs">2-1/8&quot;
                                                                                                        &nbsp;&nbsp;(54mm)</a>
                                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2-1%2F4%22" id="2-1/4-plugs">2-1/4&quot;
                                                                                                        &nbsp;&nbsp;(57mm)</a>
                                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2-1%2F2%22" id="2-1/2-plugs">2-1/2&quot;
                                                                                                        &nbsp;&nbsp;(63.5mm)</a>
                                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2-5%2F8%22" id="2-5/8-plugs">2-5/8&quot;
                                                                                                        &nbsp;&nbsp;(67mm)</a>
                                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2-3%2F4%22" id="2-3/4-plugs">2-3/4&quot;
                                                                                                        &nbsp;&nbsp;(70mm)</a>
                                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=2-7%2F8%22" id="2-7/8-plugs">2-7/8&quot;
                                                                                                        &nbsp;&nbsp;(73mm)</a>
                                                                                                <a class="dropdown-item track-gauges-dropdown" href="/products.asp?jewelry=plugs&amp;gauge=3%22" id="3-inch-plugs">3&quot;
                                                                                                        &nbsp;&nbsp;(76mm)</a>
                                                                        </div>
                                                                </div>
                                                        </div>
                                                </div>
                                        </li>
                                        <li class="nav-item dropdown position-static p-1">
                                                <a class="nav-link dropdown-toggle text-light py-2 px-3 px-lg-1 header-menu-open" href="#" id="brandsDropdown"
                                                        role="button" data-toggle="dropdown" aria-haspopup="true"
                                                        aria-expanded="false">
                                                        Brands</a>
                                                        <div class="small text-secondary px-3 d-lg-none">Browse jewelry by all our major brand names</div>
                                                <div class="dropdown-menu w-100 m-0 border-0 rounded-0 nav-dropdown"
                                                        aria-labelledby="brandsDropdown">
                                                        <div class="container-fluid pb-3">
                                                                <div class="row nav-brands">
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=alchemy+adornment" id="brand-alchemy">
                                                                                <img class="lazyload" data-src="/images/navigation/alchemy.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=atlas+glass" id="brand-atlas-glass">
                                                                                <img class="lazyload" data-src="/images/navigation/atlas.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=body+circle" id="brand-body-circle">
                                                                                <img  class="lazyload" data-src="/images/navigation/bodycircle.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=body+gems" id="brand-body-gems">
                                                                                <img class="lazyload" data-src="/images/navigation/bodygems.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=body vibe" id="brand-body-vibe">
                                                                                <img class="lazyload" data-src="/images/navigation/bodyvibe.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=buddha+jewelry" id="brand-buddha">
                                                                                <img class="lazyload" data-src="/images/navigation/buddha.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=diablo+organics" id="brand-diablo">
                                                                                <img class="lazyload" data-src="/images/navigation/diablo.png"
                                                                                        alt="Browse Diablo Organics jewelry"
                                                                                        title="Browse Diablo Organics jewelry" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=element" id="brand-element">
                                                                                <img class="lazyload" data-src="/images/navigation/element.png"
                                                                                        alt="Browse Element jewelry"
                                                                                         />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=gorilla+glass" id="brand-gorilla">
                                                                                <img class="lazyload" data-src="/images/navigation/gorilla.png"
                                                                                        title="Browse Gorilla Glass jewelry"
                                                                                        alt="Browse Gorilla Glass jewelry" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=glasswear" id="brand-glasswear">
                                                                                <img class="lazyload" data-src="/images/navigation/glasswear.png"
                                                                                        title="Browse Glasswear Studios jewelry"
                                                                                        alt="Browse Glasswear Studios jewelry" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=half+tone" id="brand-halftone">
                                                                                <img class="lazyload" data-src="/images/navigation/halftone.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=invictus" id="brand-invictus">
                                                                                <img class="lazyload" data-src="/images/navigation/invictus.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=kaos+softwear" id="brand-kaos">
                                                                                <img class="lazyload" data-src="/images/navigation/kaos.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=le+roi" id="brand-leroi">
                                                                                <img class="lazyload" data-src="/images/navigation/leroi.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=maya+organic" id="brand-maya">
                                                                                <img class="lazyload" data-src="/images/navigation/maya.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=metal+mafia" id="brand-metal-mafia">
                                                                                <img class="lazyload" data-src="/images/navigation/metalmafia.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=neometal" id="brand-neometal">
                                                                                <img class="lazyload" data-src="/images/navigation/neometal.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=norvoch" id="brand-norvoch">
                                                                                <img class="lazyload" data-src="/images/navigation/norvoch.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=oracle" id="brand-oracle">
                                                                                <img class="lazyload" data-src="/images/navigation/oracle.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=quetzalli" id="brand-quetzalli">
                                                                                <img class="lazyload" data-src="/images/navigation/quetzalli.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=sm316" id="brand-sm316">
                                                                                <img class="lazyload" data-src="/images/navigation/sm316.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=tawapa" id="brand-tawapa">
                                                                                <img class="lazyload" data-src="/images/navigation/tawapa.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=urban+star" id="brand-urban-star">
                                                                                <img class="lazyload" data-src="/images/navigation/urbanstar.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=venus+by+maria+tash" id="brand-maria-tash">
                                                                                <img class="lazyload" data-src="/images/navigation/mariatash.png" />
                                                                        </a>
                                                                        <a class="track-brands-dropdown" href="/products.asp?brand=wildcat" id="brand-wildcat">
                                                                                <img class="lazyload" data-src="/images/navigation/wildcat.png" />
                                                                        </a>
                                                                </div>
                                                        </div>
                                                </div>
                                        </li>
                                        
                                        <li class="nav-item dropdown position-static  p-1">
                                                <a class="nav-link dropdown-toggle text-light py-2 px-3 px-lg-1 header-menu-open" href="#" id="moreDropdown"
                                                        role="button" data-toggle="dropdown" aria-haspopup="true"
                                                        aria-expanded="false">
                                                        More
                                                </a>
                                                <div class="small text-secondary px-3 d-lg-none">Aftercare, tools, storage, gift certificates, &amp; BAF Gear</div>
                                                <div class="dropdown-menu w-100 m-0 border-0 rounded-0 nav-dropdown"
                                                        aria-labelledby="moreDropdown">
                                                        <div class="container-fluid pb-3">
                                                                <div class="row">
                                                                        <div class="col-xxs-12 col-sm-4 col-lg-2  mx-lg-auto">
                                                                                <a class="dropdown-item track-more-dropdown" id="recently-restocked" href="/products.asp?restock=restock">
                                                                                        <h6 class="d-inline pr-1">Recently
                                                                                                Restocked</h6>
                                                                                </a>
                                        <% if request.cookies("OrderAddonsActive") = "" then  %>
                                                                                <a class="dropdown-item track-more-dropdown" id="gift-certificates" href="/gift-certificate.asp">
                                                                                        <h6 class="d-inline pr-1">Gift
                                                                                                certificates</h6>
                                                                                </a>
                                                        <% end if %>
                                                                        </div>
                                                                        <div class="col-xxs-6 col-sm-4 col-lg-2  mx-lg-auto">
                                                                                        <a class="h6 d-block pb-1 border-bottom border-secondary track-more-dropdown" href="/products.asp?jewelry=aftercare&amp;jewelry=tools&amp;jewelry=storage&amp;jewelry-cleansers&amp;jewelry=lotion-oil" id="image-aftercare"><img data-src="/images/navigation/navi-400-aftercare.jpg"
                                                                                                class="img-fluid  img-75 d-block lazyload">
                                                                                        
                                                                                                Aftercare <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                        all</span></span>
                                                                                        </a>
                                                                                        <a class="dropdown-item track-more-dropdown" href="/products.asp?jewelry=cleansers" id="aftercare-cleansers">
                                                                                                Cleansers
                                                                                        </a>
                                                                                        <a class="dropdown-item track-more-dropdown" href="/products.asp?jewelry=lotion-oil" id="aftercare-oils">
                                                                                                Lotions &amp; oils
                                                                                        </a>
                                                                                        <a class="dropdown-item track-more-dropdown" href="/products.asp?jewelry=tools" id="aftercare-tools">
                                                                                                Tools
                                                                                        </a>
                                                                                        <a class="dropdown-item track-more-dropdown" href="/products.asp?jewelry=storage" id="aftercare-storage">
                                                                                                Storage
                                                                                        </a>
                                                                                        <a class="dropdown-item track-more-dropdown" href="/productdetails.asp?ProductID=1464" id="aftercare-sterilization">
                                                                                                Sterilization
                                                                                        </a>
                                                                        </div>
                                                                                                                                   <div class="col-xxs-6 col-sm-4 col-lg-2  mx-lg-auto">
                                                                                        <a class="h6 d-block pb-1 border-bottom border-secondary track-more-dropdown" href="/products.asp?jewelry=gear" id="image-gear"><img data-src="/images/navigation/navi-400-baf-gear.jpg"
                                                                                        class="img-fluid 
                                                                                                img-75 d-block lazyload">
                                                                               
                                                                                        BAF Gear <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                                                                                all</span></span>
                                                                                </a>
                                                                        </div>
                <div class="col-xxs-6 col-sm-4 col-lg-2  mx-lg-auto">
                        <a class="h6 d-block pb-1 border-bottom border-secondary track-more-dropdown" href="/products.asp?jewelry=accessories" id="image-accessories"><img data-src="/images/navigation/navi-400-accessories.jpg"
                                class="img-fluid 
                                        img-75 d-block lazyload">
                       
                                Accessories <span class="d-block d-xl-inline  ml-xl-4"><span class="badge badge-secondary font-weight-normal">View
                                        all</span></span>
                        </a>
                        
                </div>
                <div class="col-xxs-6 col-sm-4 col-lg-2  mx-lg-auto"></div>
                <div class="col-xxs-6 col-sm-4 col-lg-2  mx-lg-auto"></div>
                                                                </div>
                                                        </div>
                                                </div>
                                        </li>
                                        <li class="nav-item p-1 d-lg-none">
                                                <a class="nav-link py-2 px-3 px-lg-1 header-menu-link" href="/contact.asp" id="topnav-contact-us-link">
                                                        Contact Us
                                                </a>
                                        </li>
                                        <% if request.cookies("adminuser") = "yes" then %>
                                        <li class="nav-item dropdown p-1">
                                                <a class="nav-link dropdown-toggle py-2 px-3 px-lg-1" href="#" id="sandbox-menu" role="button"
                                                        data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                                                        Sandbox
                                                </a>
                                                <div class="dropdown-menu m-0 border-0 rounded-0" aria-labelledby="sandbox-menu">
                                                        <a class="dropdown-item header-menu-link" href="?inactive=yes">Show inactives</a>
                                                        <a class="dropdown-item header-menu-link" href="?inactive=no">Don't show
                                                                inactives</a>
                                                        <a href="#" class="dropdown-item header-menu-link toggle-sandbox" data-sandbox="OFF">Turn
                                                                sandbox off</a>
                                                        <a href="#" class="dropdown-item header-menu-link toggle-sandbox" data-sandbox="ON">Turn
                                                                sandbox on</a>
                                                </div>
                                        </li>
                                        <li class="nav-item pt-3">
                                                <% if session("sandbox") = "ON" then %>
                                                <a class="nav-link badge badge-success text-light">Sandbox ON</a>
                                                <% end if %>
                                                <% if session("inactive") = "yes" then %>
                                                <a class="nav-link badge badge-success text-light">Inactives showing</a>
                                                <% end if %>
                                        </li>
                                        <% end if ' logged in as admin user %>
                                </ul>
                        </div>
                </div>
        </nav>
        <!-- Sign In Modal -->
        <div class="modal fade" id="signin" tabindex="-1" role="dialog" aria-labelledby="signinLabel" aria-hidden="true">
                <div class="modal-dialog" role="document">
                        <div class="modal-content">
                                <div class="modal-header">
                                        <h5 class="modal-title" id="signinLabel">Sign In</h5>
                                        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                                                <span aria-hidden="true">&times;</span>
                                        </button>
                                </div>
                                <div class="modal-body">
                                        <form class="needs-validation" id="frm-signin" novalidate>
												<div class="form-group">
													<div id="google_sign_in"></div>
												</div>
												<div style="width: 100%;height: 13px;border-bottom: 1px solid #c7c7c7;text-align: center;margin-bottom: 25px;">
													<div style="color: #969191; font-size: 15px; background-color: #fff; margin:auto auto; width: 40px;"> OR </div>
												</div>
                                                <div class="form-group">
                                                        <input type="email" name="email" class="form-control"
                                                                placeholder="E-mail" required>
                                                        <div class="invalid-feedback">
                                                                Please enter a valid e-mail address
                                                        </div>
                                                </div>
                                                <div class="form-group">
                                                        <input type="password" name="password" class="form-control"
                                                                placeholder="Password" required>
                                                        <div class="invalid-feedback">
                                                                Password is required
                                                        </div>
                                                </div>
                                                <div class="alert alert-danger alert-dismissible collapse alert-signin"
                                                        role="alert">
                                                        <span class="signin-message"></span>
                                                        <button type="button" class="close" data-hide="alert"
                                                                aria-label="Close">
                                                                <span aria-hidden="true">&times;</span>
                                                        </button>
                                                </div>
                                                <div class="text-center">
                                                        <button type="submit" class="btn btn-block btn-purple"
                                                                id="btn_signin" name="btn_signin">Sign In</button>
                                                </div>
                                        </form>
                                </div>
                                <h6 class="text-muted text-center mt-2 mb-2">Don't have an account yet?</h6>
                                <div class="d-block text-center">
                                        <button class="btn btn-purple btn-sm" data-toggle="modal" data-target="#createaccount"
                                                data-dismiss="modal">Create a new account</button>
                                        <span class="small d-block p-4">
                                                <a href="" data-toggle="modal" data-target="#forgotpassword"
                                                        data-dismiss="modal" href="#">Forgot your password?</a>
                                        </span>
                                </div>
                        </div>
                </div>
        </div>
        <!-- Forgot password Modal -->
        <div class="modal fade" id="forgotpassword" tabindex="-1" role="dialog" aria-labelledby="forgotpasswordLabel"
                aria-hidden="true">
                <div class="modal-dialog" role="document">
                        <div class="modal-content">
                                <div class="modal-header">
                                        <h5 class="modal-title" id="forgotpasswordLabel">Retrieve your password</h5>
                                        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                                                <span aria-hidden="true">&times;</span>
                                        </button>
                                </div>
                                <div class="modal-body">
                                        <form class="needs-validation" name="frmForgotPass" id="frmForgotPass" novalidate>
                                                <div class="form-group">
                                                        <input class="form-control" type="email" name="email"                         id="forgot_email"                                        placeholder="E-mail" required>
                                                                <div class="invalid-feedback">
                                                                                Please enter a valid e-mail address
                                                                        </div>
                                                </div>
                                                <div id="message-forgot"></div>
                                                <div class="text-center">
                                                        <button type="submit" name="btn-forgot" class="btn btn-block btn-purple">Send
                                                                e-mail</button>
                                                </div>
                                        </form>
                                        <div class="small d-block pt-4 text-center">
                                                        <a href="" data-toggle="modal" data-target="#signin"
                                                                data-dismiss="modal" href="#">Sign in</a>
                                                      
                                                        <span class="mx-2">|</span>
                                                          <a href="" data-toggle="modal" data-target="#createaccount"
                                                          data-dismiss="modal" href="#">Create account</a>
                                                  </div>
                                </div>
                        </div>
                </div>
        </div>
        </div>
        
        <!-- Create new account Modal -->
        <div class="modal fade" id="createaccount" tabindex="-1" role="dialog" aria-labelledby="createaccountLabel"
                aria-hidden="true">
                <div class="modal-dialog" role="document">
                        <div class="modal-content">
                                <div class="modal-header">
                                        <h5 class="modal-title" id="createaccountLabel">Create a new account</h5>
                                        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                                                <span aria-hidden="true">&times;</span>
                                        </button>
                                </div>
                                <form class="needs-validation" name="frm-register" id="frm-register" novalidate>
                                        <div class="modal-body">
                                                <div class="alert alert-info small p-1">
                                                <div class="alert-link">Account benefits:</div>
                                                        <ul class="mb-0 pb-0">
                                                                <li>Earn points for jewelry reviews & photo submissions</li>
                                                                <li>Save your cart & store items for later</li>
                                                                <li>Wishlists</li>
                                                                <li>Saved searches & filters</li>
                                                        </ul>
                                                </div>
                                                <div class="form-group">
                                                        <label for="regEmail">E-mail <span class="text-danger">*</span></label>
                                                        <input class="form-control" name="e-mail" type="email"
                                                                id="regEmail" size="30" required/>
                                                        <div class="invalid-feedback">
                                                                Please enter a valid e-mail address
                                                        </div>
                                                </div>
                                                <div class="form-group">
                                                        <label for="password_confirmation">Password
                                                                <span class="text-danger">*</span></label>
                                                        <input class="form-control"  name="password_confirmation" id="password_confirmation"
                                                                type="password" size="30" required />
                                                                <div class="invalid-feedback">
                                                                                Password is required
                                                                        </div>
                                                </div>
                                                <div class="form-group">
                                                        <label for="password">Re-type password
                                                                <span class="text-danger">*</span></label>
                                                        <input  class="form-control" name="password" id="Regpassword" type="password" size="30"
                                                         required />
                                                         <div class="invalid-feedback">
                                                                        Password is required
                                                                </div>
                                                </div>
                                                <input type="hidden" name="status" value="register" />
                                                <input type="hidden" name="check" value="" />
                                                <div id="message-create-account"></div>
                                        </div>
                                        <div class="modal-footer">
                                                <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
                                                <button type="submit" class="btn btn-purple modal-submit" id="btn-create-account">Create</button>
                                        </div>
                                        <div class="small d-block py-2 text-center">
                                                        <a href="" data-toggle="modal" data-target="#signin"
                                                                data-dismiss="modal" href="#">Already have an account? Sign in here</a>
                                                        </div>
                                </form>
                        </div>
                </div>
        </div>
        </div>