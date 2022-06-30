<!--#include virtual="/Connections/klaviyo.asp" -->
<!--#include virtual="/Connections/google-oauth-credentials.inc" -->
<!-- Installed April 2021 - Global site tag (gtag.js) - Google Analytics -->
<!-- Google Tag Manager -->
<script>(function(w,d,s,l,i){w[l]=w[l]||[];w[l].push({'gtm.start':
        new Date().getTime(),event:'gtm.js'});var f=d.getElementsByTagName(s)[0],
        j=d.createElement(s),dl=l!='dataLayer'?'&l='+l:'';j.async=true;j.src=
        'https://www.googletagmanager.com/gtm.js?id='+i+dl;f.parentNode.insertBefore(j,f);
        })(window,document,'script','dataLayer','GTM-W46N98J');</script>
        <!-- GTM-W46N98J is for Google Tag Manager -->
        <!-- End Google Tag Manager -->
<!-- Installed April 2021 - Global site tag (gtag.js) - Google Analytics -->
<!-- G-CG6EYC3NFB is GA4 Measurement ID located under Google Analytics > Admin > Property > Data Stream -->
<!-- UA-32113869-1 is the standard Universal Analytics (old style) Tracking ID -->
<script async src="https://www.googletagmanager.com/gtag/js?id=G-CG6EYC3NFB"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());

  gtag('config', 'G-CG6EYC3NFB');
</script>
<script>(function(d){var e=d.createElement('script');e.src='https://td.yieldify.com/yieldify/code.js?w_uuid=ef4e975b-07ee-476e-86fd-1f9ae619a60f&k=1&loca='+window.location.href;e.async=true;d.getElementsByTagName('head')[0].appendChild(e);}(document));</script>
<script
  async type="text/javascript"
  src="//static.klaviyo.com/onsite/js/klaviyo.js?company_id=<%=klaviyo_public_key%>"
></script>
<!-- BEGIN TIK TOK -->
<script> !function (w, d, t) {   w.TiktokAnalyticsObject=t;var ttq=w[t]=w[t]||[];ttq.methods=["page","track","identify","instances","debug","on","off","once","ready","alias","group","enableCookie","disableCookie"],ttq.setAndDefer=function(t,e){t[e]=function(){t.push([e].concat(Array.prototype.slice.call(arguments,0)))}};for(var i=0;i<ttq.methods.length;i++)ttq.setAndDefer(ttq,ttq.methods[i]);ttq.instance=function(t){for(var e=ttq._i[t]||[],n=0;n<ttq.methods.length;n++)ttq.setAndDefer(e,ttq.methods[n]);return e},ttq.load=function(e,n){var i="https://analytics.tiktok.com/i18n/pixel/events.js";ttq._i=ttq._i||{},ttq._i[e]=[],ttq._i[e]._u=i,ttq._t=ttq._t||{},ttq._t[e]=+new Date,ttq._o=ttq._o||{},ttq._o[e]=n||{};var o=document.createElement("script");o.type="text/javascript",o.async=!0,o.src=i+"?sdkid="+e+"&lib="+t;var a=document.getElementsByTagName("script")[0];a.parentNode.insertBefore(o,a)};   ttq.load('C6EGBPOA2TFR2CRB1GLG');   ttq.page(); }(window, document, 'ttq'); </script>
<!-- END TIK TOK -->
<!-- Reddit Pixel -->
<script>
  !function(w,d){if(!w.rdt){var p=w.rdt=function(){p.sendEvent?p.sendEvent.apply(p,arguments):p.callQueue.push(arguments)};p.callQueue=[];var t=d.createElement("script");t.src="https://www.redditstatic.com/ads/pixel.js",t.async=!0;var s=d.getElementsByTagName("script")[0];s.parentNode.insertBefore(t,s)}}(window,document);rdt('init','t2_8g6bx251', {"optOut":false,"useDecimalCurrencyValues":true,"aaid":"G-CG6EYC3NFB"});rdt('track', 'PageVisit');
  </script>
   <!-- End Reddit Pixel -->

  <!-- Pinterest Tag -->
<script>
  !function(e){if(!window.pintrk){window.pintrk = function () {
  window.pintrk.queue.push(Array.prototype.slice.call(arguments))};var
    n=window.pintrk;n.queue=[],n.version="3.0";var
    t=document.createElement("script");t.async=!0,t.src=e;var
    r=document.getElementsByTagName("script")[0];
    r.parentNode.insertBefore(t,r)}}("https://s.pinimg.com/ct/core.js");
  pintrk('load', '2614209404774', {em: '<user_email_address>'});
  pintrk('page');
  </script>
  <noscript>
  <img height="1" width="1" style="display:none;" alt=""
    src="https://ct.pinterest.com/v3/?event=init&tid=2614209404774&pd[em]=<hashed_email_address>&noscript=1" />
  </noscript>
  <!-- end Pinterest Tag -->
<% If not rsGetUser.EOF and request.cookies("ID") <> "" then %>
<script>
    var _learnq = _learnq || [];
    _learnq.push(['identify', {
      '$email' : '<%= rsGetUser("email") %>',
      '$first_name' : '<%= rsGetUser("customer_first") %>',
      '$last_name' : '<%= rsGetUser("customer_last") %>'
    }]);
</script>
<% end if %>



        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="facebook-domain-verification" content="muhfd0zbw04kmod638n0srrqinj67i" />
        <meta name="description" content="<%= page_description %>">
        <% if var_extra_head_inc = "homepage" then %>
        <meta name="google-site-verification" content="WAeYp74lwrgsm6m1dfjFwmtRzL3MRidv2O5qBt2k7Wg" />
        <% end if %>
        <title><%= page_title %></title>
        <link rel="apple-touch-icon" sizes="180x180" href="/apple-touch-icon.png">
        <link rel="icon" type="image/png" sizes="32x32" href="/favicon-32x32.png">
        <link rel="icon" type="image/png" sizes="192x192" href="/android-chrome-192x192.png">
        <link rel="icon" type="image/png" sizes="256x256" href="/android-chrome-256x256.png">
        <link rel="icon" type="image/png" sizes="384x384" href="/android-chrome-384x384.png">
        <link rel="icon" type="image/png" sizes="512x512" href="/android-chrome-512x512.png">
        <link rel="icon" type="image/png" sizes="16x16" href="/favicon-16x16.png">
        <meta name="google-signin-client_id" content="<%=google_oauth_clientId%>.apps.googleusercontent.com">
        <link rel="manifest" href="/webmanifest.json">
        <link rel="mask-icon" href="/safari-pinned-tab.svg" color="#5bbad5">
        <link rel="shortcut icon" href="/favicon.ico">
        <meta name="msapplication-TileColor" content="#2b5797">
        <meta name="theme-color" content="#ffffff">
        <% if request.cookies("darkmode") <> "on" then 
        var_applepay_color = "black"
        %> 
        <link href="/CSS/baf.min.css?v=042822" id="lightmode" rel="stylesheet" type="text/css" />
        <% else 
        var_applepay_color = "white"
        %>
        <link href="/CSS/baf-dark.min.css?v=042822" id="darkmode" rel="stylesheet" type="text/css" />
        <% end if %>
        <link href="/CSS/ion.rangeslider.min.css?v=061121" rel="stylesheet" type="text/css" />
        <link href="/CSS/media-max768.min.css" media="screen and (max-width: 768px)" rel="stylesheet" type="text/css" />
        <link href="/CSS/media-max1024.min.css" media="screen and (max-width: 1024px)" rel="stylesheet" type="text/css" />
        <link href="/CSS/media-min992-max1024.min.css" media="screen and (min-width: 992px) and (max-width: 1024px)" rel="stylesheet" type="text/css" />
        <link href="/CSS/media-min-1025.min.css" media="screen and (min-width: 1025px)" rel="stylesheet" type="text/css" />
        <link href="/CSS/media-min-1600.min.css?v=112818" media="screen and (min-width: 1600px)" rel="stylesheet" type="text/css" />
        <link href="/CSS/fortawesome/css/external-min.css?v=031920" rel="stylesheet" type="text/css" />