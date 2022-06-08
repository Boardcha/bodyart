<% if request.querystring("pagenumber") = "" or request.querystring("pagenumber") = "1" then 

if instr(request.querystring("jewelry"), ",") > 0 then
	landing_multiple = "yes"
end if

var_url_querystring = "?" & Request.ServerVariables("QUERY_STRING")


var_gauge = request.querystring("gauge")
var_material_linkchange = Server.HTMLEncode(request.querystring("material"))
var_flare_linkchange = Server.HTMLEncode(request.querystring("flare_type"))
var_keywords_linkchange = Server.HTMLEncode(request.querystring("keywords"))
var_jewelry_linkchange = Server.HTMLEncode(request.querystring("jewelry"))

if request.querystring("brand") <> "" then
	var_brand_linkchange = "&amp;brand=" & Server.HTMLEncode(request.querystring("brand"))
end if
if request.querystring("gauge") <> "" then
	var_gauge_linkchange = "&amp;gauge=" & Server.HTMLEncode(request.querystring("gauge"))
end if

if var_flare_linkchange <> "" then 
	var_flare_linkchange =  "&flare_type=" & var_flare_linkchange
end if
if var_material_linkchange <> "" then 
	var_material_linkchange =  "&material=" & var_material_linkchange
end if
if var_keywords_linkchange <> "" then 
	var_keywords_linkchange = "&keywords=" & var_keywords_linkchange
end if

if request.querystring("jewelry") = "plugs" or request.querystring("jewelry") = "finger-ring" or request.querystring("jewelry") = "septum" or instr(request.querystring("jewelry"),"septum") > 0 or request.querystring("piercing") = "septum" or request.querystring("jewelry") = "balls" or request.querystring("jewelry") = "nose-ring" or request.querystring("jewelry") = "nose-hoop" or request.querystring("jewelry") = "nose-threadless" or instr(request.querystring("jewelry"), "belly") > 0 or request.querystring("jewelry") = "labret" or request.querystring("jewelry") = "barbell" or request.querystring("jewelry") = "captive" or instr(request.querystring("jewelry"), "balls") > 0 then

if var_gauge = "8g" OR var_gauge = "6g" OR var_gauge = "4g" OR var_gauge = "2g" OR var_gauge = "0g" OR InStr(var_gauge_linkchange, "00g") > 0 OR InStr(var_gauge_linkchange, "7/16") > 0 OR var_gauge = "1/2""" OR var_gauge = "9/16""" OR var_gauge = "5/8""" OR var_gauge = "3/4""" OR var_gauge = "7/8""" OR InStr(var_gauge_linkchange, "1&quot;") > 0 OR InStr(var_gauge_linkchange, "1-1/8") > 0 OR InStr(var_gauge_linkchange, "1-1/4") > 0 OR InStr(var_gauge_linkchange, "1-3/8") > 0 OR InStr(var_gauge_linkchange, "1-1/2") > 0 then
	show_plugs_landing = "yes"
end if
%>

<% if instr(request.querystring("jewelry"), "plugs") > 0 and (landing_multiple <> "yes" or show_plugs_landing = "yes") then 
var_plugs_string = "?jewelry=plugs" & var_gauge_linkchange & var_brand_linkchange & "&amp;flare_type=" & var_flare_linkchange & "&amp;material=" & var_material_linkchange & "&amp;keywords=" & var_keywords_linkchange & "&amp;"

%>

<div class="card mt-3 mb-1">
	<div class="card-header p-2">
			<h5 class="p-0 m-0">Popular Categories</h5>
	</div>
	<div class="card-body p-2">
<div class="d-flex flex-row flex-wrap mb-3h">
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 plugs-landing-link" id="under20-<%= request.querystring("gauge") %>" href="?jewelry=plugs<%= var_gauge_linkchange %><%= var_brand_linkchange %><%= var_material_linkchange %><%= var_keywords_linkchange %>&price=0%3B20">Plugs under $20</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 plugs-landing-link" id="double-flare-<%= request.querystring("gauge") %>" href="?jewelry=plugs<%= var_gauge_linkchange %><%= var_brand_linkchange %><%= var_material_linkchange %><%= var_keywords_linkchange %>&flare_type=Double+flare">Double flare</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 plugs-landing-link" id="single-flare-<%= request.querystring("gauge") %>" href="?jewelry=plugs<%= var_gauge_linkchange %><%= var_brand_linkchange %><%= var_material_linkchange %><%= var_keywords_linkchange %>&flare_type=Single+flare">Single flare</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 plugs-landing-link" id="no-flare-<%= request.querystring("gauge") %>" href="?jewelry=plugs<%= var_gauge_linkchange %><%= var_brand_linkchange %><%= var_material_linkchange %><%= var_keywords_linkchange %>&flare_type=No+flare">No flares</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 plugs-landing-link" id="tunnels-<%= request.querystring("gauge") %>" href="?jewelry=plugs<%= var_flare_linkchange %><%= var_gauge_linkchange %><%= var_brand_linkchange %><%= var_material_linkchange %>&keywords=eyelet+tunnel+tunnels">Tunnels &amp; eyelets</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 plugs-landing-link" id="teardrops-<%= request.querystring("gauge") %>" href="?jewelry=plugs<%= var_flare_linkchange %><%= var_gauge_linkchange %><%= var_brand_linkchange %><%= var_material_linkchange %>&keywords=teardrop">Teardrops</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 plugs-landing-link" id="saddle-<%= request.querystring("gauge") %>" href="?jewelry=saddle<%= var_gauge_linkchange %>">Saddles</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 plugs-landing-link" id="glass-<%= request.querystring("gauge") %>" href="?jewelry=plugs<%= var_flare_linkchange %><%= var_gauge_linkchange %><%= var_brand_linkchange %><%= var_keywords_linkchange %>&material=Glass">Glass</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 plugs-landing-link" id="silicone-<%= request.querystring("gauge") %>" href="?jewelry=plugs<%= var_flare_linkchange %><%= var_gauge_linkchange %><%= var_brand_linkchange %><%= var_keywords_linkchange %>&material=Silicone">Silicone</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 plugs-landing-link" id="organic-<%= request.querystring("gauge") %>" href="?jewelry=plugs<%= var_flare_linkchange %><%= var_gauge_linkchange %><%= var_brand_linkchange %><%= var_keywords_linkchange %>&material=Organics">Organic</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 plugs-landing-link" id="stretching-<%= request.querystring("gauge") %>" href="?jewelry=tapers<%= var_gauge_linkchange %><%= var_brand_linkchange %>">Stretching &amp; tapers</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 plugs-landing-link" id="orings-<%= request.querystring("gauge") %>" href="?jewelry=orings">O-rings</a> 


</div>
</div><!-- card body -->
</div><!-- card -->
<% end if %>

<% if request.querystring("jewelry") = "finger-ring" and landing_multiple <> "yes" then 
var_ring_string = "?jewelry=finger-ring&amp;gauge="
landing_found = 1
%>
<div class="card mt-3 mb-1">
	<div class="card-header p-2">
			<h5 class="p-0 m-0">Popular Categories</h5>
	</div>
	<div class="card-body p-2">
<div class="d-flex flex-row flex-wrap mb-3h">
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-5" href="<%= var_ring_string %>Size+5">Size 5</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-5.5" id="size-" href="<%= var_ring_string %>Size+5.5">Size 5.5</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-6" href="<%= var_ring_string %>Size+6">Size 6</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-6.5" href="<%= var_ring_string %>Size+6.5">Size 6.5</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-7" href="<%= var_ring_string %>Size+7">Size 7</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-7.8" href="<%= var_ring_string %>Size+7.5">Size 7.5</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-8" href="<%= var_ring_string %>Size+8">Size 8</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-8.5" href="<%= var_ring_string %>Size+8.5">Size 8.5</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-9" href="<%= var_ring_string %>Size+9">Size 9</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-9.5" href="<%= var_ring_string %>Size+9.5">Size 9.5</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-10" href="<%= var_ring_string %>Size+10">Size 10</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-10.5" href="<%= var_ring_string %>Size+10.5">Size 10.5</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-11" href="<%= var_ring_string %>Size+11">Size 11</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 rings-landing-link" id="size-11.5" href="<%= var_ring_string %>Size+11.5">Size 11.5</a>
</div>
</div><!-- card body -->
</div><!-- card -->
<% end if %>

<% if (instr(request.querystring("jewelry"), "septum") > 0 or instr(request.querystring("piercing"), "septum") > 0) then 


var_septum_string1 = "?jewelry=septum" & var_gauge_linkchange & var_brand_linkchange & "&amp;"
var_septum_string2 = var_gauge_linkchange & var_brand_linkchange
landing_found = 1
%>
<div class="card mt-3 mb-1">
	<div class="card-header p-2">
			<h5 class="p-0 m-0">Popular Categories</h5>
	</div>
	<div class="card-body p-2">
<div class="d-flex flex-row flex-wrap mb-3">
		<a class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 septum-landing-links" id="landing-septum-clickers"  href="<%= var_septum_string1 %>keywords=clicker">Clickers</a>
	<a class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 septum-landing-links" id="landing-septum-basics" href="<%= var_septum_string1 %>jewelry=basics">Basics</a>
	<a class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 septum-landing-links" id="landing-septum-captives"  href="?jewelry=septum-captive<%= var_septum_string2 %>">Captives</a>
	<a class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 septum-landing-links" id="landing-septum-pinchers"  href="<%= var_septum_string1 %>keywords=pincher">Pinchers</a>
	<a class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 septum-landing-links" id="landing-septum-tusks"  href="?jewelry=septum-spike<%= var_septum_string2 %>">Tusks &amp; Spikes</a>
	<a class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 septum-landing-links" id="landing-septum-circulars"  href="?jewelry=circular<%= var_septum_string2 %>">Circulars</a>
	<a class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 septum-landing-links" id="landing-septum-seamless"  href="<%= var_septum_string1 %>keywords=seamless">Seamless</a>
	<a class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 septum-landing-links" id="landing-septum-gold"  href="<%= var_septum_string1 %>material=gold">Gold</a>
	<a class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 septum-landing-links" id="landing-septum-retainers"  href="?keywords=retainer&amp;piercing=septum<%= var_septum_string2 %>">Retainers</a>
	<a class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 septum-landing-links" id="landing-septum-plugs"  href="?keywords=plug&amp;piercing=septum<%= var_septum_string2 %>">Plugs</a>
</div>
</div><!-- card body -->
</div><!-- card -->
<% end if %>

<%
if (instr(request.querystring("jewelry"), "balls") > 0 or instr(request.querystring("jewelry"), "beads") > 0) then 


var_ends_string = "?jewelry=balls" & var_gauge_linkchange & var_brand_linkchange & "&amp;"
landing_found = 1
%>
<div class="card mt-3 mb-1">
	<div class="card-header p-2">
			<h5 class="p-0 m-0">Popular Categories</h5>
	</div>
	<div class="card-body p-2">
<div class="d-flex flex-row flex-wrap mb-3h">
		<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 balls-landing-links" id="landing-balls-internal" href="<%= var_ends_string %>threading=internally threaded">Internal</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 balls-landing-links" id="landing-balls-threadless" href="<%= var_ends_string %>threading=threadless">Threadless</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 balls-landing-links" id="landing-balls-gold" href="<%= var_ends_string %>material=gold">Gold</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 balls-landing-links" id="landing-balls-external" href="<%= var_ends_string %>threading=externally threaded">External</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 balls-landing-links" id="landing-balls-basics" href="<%= var_ends_string %>jewelry=basics">Basics</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 balls-landing-links" id="landing-balls-beads" href="?jewelry=beads">Beads</a>
</div>
</div><!-- card body -->
</div><!-- card -->
<% end if %>

<% if instr(request.querystring("jewelry"), "nose") > 0 and landing_multiple <> "yes" then 

var_nose_string = "jewelry=nose-ring" & var_gauge_linkchange & var_brand_linkchange & var_keywords_linkchange & var_material_linkchange & "&amp;"
landing_found = 1

%>
<div class="card mt-3 mb-1">
	<div class="card-header p-2">
			<h5 class="p-0 m-0">Popular Categories</h5>
	</div>
	<div class="card-body p-2">
<div class="d-flex flex-row flex-wrap mb-3h">
		<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 nose-landing-links" id="landing-nosescrews" href="?<%= replace(var_nose_string, var_keywords_linkchange, "") %>keywords=nosescrews">Nose screws</a>
		<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 nose-landing-links" id="landing-nose-rings-hoops" href="?<%= replace(replace(var_nose_string, "jewelry=nose-ring&amp;", ""), var_keywords_linkchange, "") %>jewelry=nose-hoop">Rings & Hoops</a>
		<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 nose-landing-links" id="landing-nose-18g"  href="?<%= replace(var_nose_string, var_gauge_linkchange, "") %>gauge=18g">18 Gauge</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 nose-landing-links" id="landing-nose-20g"  href="?<%= replace(var_nose_string, var_gauge_linkchange, "") %>gauge=20g">20 Gauge</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 nose-landing-links" id="landing-nose-threadless" href="?<%= replace(replace(var_nose_string, "jewelry=nose-ring&amp;", ""), var_keywords_linkchange, "") %>jewelry=nose-threadless">Threadless</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 nose-landing-links" id="landing-nose-threadless" href="?<%= var_nose_string %>material=gold">Gold</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 nose-landing-links" id="landing-nose-basics"  href="?<%= replace(var_nose_string, var_keywords_linkchange, "") %>jewelry=basics">Basics</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2" id="landing-nosebones" href="?<%= replace(var_nose_string, var_keywords_linkchange, "") %>keywords=nosebones">Nose studs</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 nose-landing-links" id="landing-nose-retainers" href="?<%= replace(var_nose_string, var_keywords_linkchange, "") %>keywords=retainer">Retainers</a>
	



</div>
</div><!-- card body -->
</div><!-- card -->
<% end if %>

<% if (instr(request.querystring("jewelry"), "belly") > 0 or instr(request.querystring("jewelry"), "navel") > 0) and landing_multiple <> "yes" then 

var_navel_string = "jewelry=belly" & var_gauge_linkchange & var_brand_linkchange & "&amp;"
landing_found = 1
%>
<div class="card mt-3 mb-1">
	<div class="card-header p-2">
			<h5 class="p-0 m-0">Popular Categories</h5>
	</div>
	<div class="card-body p-2">
<div class="d-flex flex-row flex-wrap mb-3h">
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 belly-landing-links" id="landing-belly-no-dangles" href="?jewelry=belly-simple">No dangles</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 belly-landing-links" id="landing-belly-with-dangles" href="?jewelry=belly-dangle">With dangles</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 belly-landing-links" id="landing-belly-retainers" href="?<%= var_navel_string %>keywords=retainer">Retainers</a>
</div>
</div><!-- card body -->
</div><!-- card -->
<% end if %>

<% if (instr(request.querystring("jewelry"), "labret") > 0 or instr(request.querystring("piercing"), "labret") > 0) and landing_multiple <> "yes" then 

var_labret_string = "?jewelry=labret" & var_gauge_linkchange & var_brand_linkchange & "&amp;"
landing_found = 1
%>
<div class="card mt-3 mb-1">
	<div class="card-header p-2">
			<h5 class="p-0 m-0">Popular Categories</h5>
	</div>
	<div class="card-body p-2">
<div class="d-flex flex-row flex-wrap mb-3h">
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 labret-landing-links" id="landing-labret-basics" href="<%= var_labret_string %>jewelry=basics">Basics</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 labret-landing-links" id="landing-labret-design-ends" href="?jewelry=labret-design">Design ends</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 labret-landing-links" id="landing-labret-threadless" href="<%= var_labret_string %>threading=threadless">Threadless</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 labret-landing-links" id="landing-labret-captives" href="?piercing=labret&amp;jewelry=captive">Rings</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 labret-landing-links" id="landing-labret-stretched" href="?jewelry=labret-stretched">Large labrets</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 labret-landing-links" id="landing-labret-retainers" href="<%= var_labret_string %>keywords=retainer">Retainers</a>
</div>
</div><!-- card body -->
</div><!-- card -->
<% end if %>

<% if instr(request.querystring("jewelry"), "barbell") > 0 and landing_multiple <> "yes" then 

var_barbell_string = "?jewelry=barbell" & var_gauge_linkchange & var_brand_linkchange & "&amp;"
landing_found = 1
%>
<div class="card mt-3 mb-1">
	<div class="card-header p-2">
			<h5 class="p-0 m-0">Popular Categories</h5>
	</div>
	<div class="card-body p-2">
<div class="d-flex flex-row flex-wrap mb-3h">
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 barbell-landing-links" id="landing-barbell-basics" href="<%= var_barbell_string %>jewelry=basics">Basics</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 barbell-landing-links" id="landing-barbell-helix" href="<%= var_barbell_string %>keywords=helix">Cartilage / Helix</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 barbell-landing-links" id="landing-barbell-industrial" href="<%= var_barbell_string %>keywords=industrial">Industrials</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 barbell-landing-links" id="landing-barbell-nipple" href="<%= var_barbell_string %>piercing=nipple">Nipple</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 barbell-landing-links" id="landing-barbell-tongue" href="<%= var_barbell_string %>keywords=tongue">Tongue</a>
</div>
</div><!-- card body -->
</div><!-- card -->
<% end if %>

<% if instr(request.querystring("jewelry"), "captive") > 0 and instr(request.querystring("jewelry"), "septum") <= 0 and landing_multiple <> "yes" then 

var_captive_string = "?jewelry=captive" & var_gauge_linkchange & var_brand_linkchange & "&amp;"
landing_found = 1
%>
<div class="card mt-3 mb-1">
	<div class="card-header p-2">
			<h5 class="p-0 m-0">Popular Categories</h5>
	</div>
	<div class="card-body p-2">
<div class="d-flex flex-row flex-wrap mb-3h">
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 captives-landing-links" id="landing-captive-basics" href="<%= var_captive_string %>jewelry=basics">Basics</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 captives-landing-links" id="landing-captive-clickers" href="<%= var_captive_string %>keywords=clicker">Clickers</a>
	<a  class="text-light bg-lightpurple p-2 mx-1 d-block mt-2 captives-landing-links" id="landing-captive-seamless" href="<%= var_captive_string %>keywords=seamless">Seamless</a>
</div>
</div><!-- card body -->
</div><!-- card -->
<% end if %>


<% 
end if ' only show card is certain jewelry categories are there
end if ' if page number = 1
%>

