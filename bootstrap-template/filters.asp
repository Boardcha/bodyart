<main>
  <div class="container-fluid">
    <div class="row">
      <div class="collapse col-12 col-lg-3 col-xl-2 px-0" id="filters"> 
        <div class="row m-0 pt-2 d-lg-none">
          <div class="col-8 h4 mt-2">Advanced search & filters</div>
          <div class="col-4 text-right">
            <button class="btn btn-outline-danger btn-sm m-1 mr-3 d-lg-none" data-toggle="collapse" data-target="#filters" type="button">Close <i class="fa fa-times"></i></button> 
          </div>
        </div>
        <form action="/products.asp" method="get" id="form-filters">
        <div class="form-group pt-2 pl-2 pr-2 m-0">
        <input class="form-control border-secondary" name="keywords" id="filter-keywords" type="search" placeholder="Enter keywords" value="<%=  Server.HTMLEncode(Sanitize(request.querystring("keywords"))) %>"  /></span>

<input type="submit" class="btn btn-purple w-100 my-1 d-none d-md-block" value="Search &amp; Apply Filters">
</div>
          <div class="d-none">
            <!-- hidden fields for sub categories and other category tags not displayed in visible list -->
            <input type="checkbox" name="jewelry" value="basics">
            <input type="checkbox" name="jewelry" value="septum-captive">
            <input type="checkbox" name="jewelry" value="septum-spike">
            <input type="checkbox" name="jewelry" value="weight-light">
            <input type="checkbox" name="jewelry" value="weight-medium">
            <input type="checkbox" name="jewelry" value="weight-heavy">
            <input type="checkbox" name="jewelry" value="weight-super-heavy">
            <input type="checkbox" name="jewelry" value="nose-threadless">
            <input type="checkbox" name="jewelry" value="nose-hoop">
            <input type="checkbox" name="jewelry" value="belly-simple">
            <input type="checkbox" name="jewelry" value="belly-dangle">
            <input type="checkbox" name="jewelry" value="labret-basic">
            <input type="checkbox" name="jewelry" value="labret-stretched">
            <input type="checkbox" name="jewelry" value="captive-cbr">
            <input type="checkbox" name="jewelry" value="nipple-shield">
            <input type="checkbox" name="jewelry" value="nipple-stirrup">
            <input type="checkbox" name="jewelry" value="nipple-capcir">
            <input type="checkbox" name="jewelry" value="halloween">
            <input type="checkbox" name="jewelry" value="earring-huggies">
            <input type="text" name="filter-stock" id="filter-stock" value="<%= request.querystring("filter-stock") %>">
            <input type="text" name="new" id="filter_new" value="<%= request.querystring("new") %>">
            <input type="text" name="restock" id="filter_restock" value="<%= request.querystring("restock") %>">
          </div>
          <% if request.querystring() <> "" then %>

              <div class="bg-secondary text-white px-2 py-1 mt-3">Current Filters
                </div>
                <div class="px-2">
              <!--#include virtual="/products/inc-selected-filters.asp"-->
             
            </div>
            <% end if %>


 <div class="bg-secondary text-white px-2 py-1 mt-3">More Filters:</div>         
          <div id="accordion" class="filters-accordion">
            <div class="card rounded-0 border-0">
              <a class="card-header collapsed h6 filter-dropdown" id="categories-head" data-toggle="collapse" data-target="#categories" aria-expanded="false" aria-controls="categories"
                href="#"><i class="fa" aria-hidden="true"></i>
                Categories
              </a>
              <div id="categories" class="collapse" aria-labelledby="categories-head" data-parent="#accordion">
                <div class="card-body filter-scroll">
                  <div class="h5 mt-3 mb-0 pb-1 w-75 border-bottom">
                    Body Jewelry
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="balls" id="filter-ends" data-friendly="Balls &amp; ends">
                    <label class="form-check-label d-block" for="filter-ends">Balls/Ends</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="beads" id="filter-beads" data-friendly="Replacement Captive Beads">
                    <label class="form-check-label d-block" for="filter-beads">Replacement Captive Beads</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="belly" id="filter-belly" data-friendly="Belly">
                    <label class="form-check-label d-block" for="filter-belly">Belly Rings</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="captive" id="filter-captives" data-friendly="Captive rings">
                    <label class="form-check-label d-block" for="filter-captives">Captive Rings</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="circular" id="filter-circular" data-friendly="Circular barbells">
                    <label class="form-check-label d-block" for="filter-circular">Circular Barbells</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="clicker" id="filter-clicker" data-friendly="Clickers">
                    <label class="form-check-label d-block" for="filter-clicker">Clickers</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="curved" id="filter-curves" data-friendly="Curved barbells">
                    <label class="form-check-label d-block" for="filter-curves">Curved Barbells</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="labret" id="filter-labrets" data-friendly="Labret jewelry">
                    <label class="form-check-label d-block" for="filter-labrets">Labrets</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="nipple" id="filter-nipple" data-friendly="Nipple jewelry">
                    <label class="form-check-label d-block" for="filter-nipple">Nipple</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="nose-ring" id="filter-nose" data-friendly="Nose jewelry">
                    <label class="form-check-label d-block" for="filter-nose">Nose / Nostril</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="chains-short" id="filter-chains-short" data-friendly="Short chains">
                    <label class="form-check-label d-block" for="filter-chains-short">Nose & ear chains</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="orings" id="filter-orings" data-friendly="O-rings">
                    <label class="form-check-label d-block" for="filter-orings">O-rings</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="pinchers" id="filter-pinchers" data-friendly="Pinchers">
                    <label class="form-check-label d-block" for="filter-pinchers">Pinchers</label>
                  </div>
                  <div class="form-check">
                  
                    <input class="form-check-input"  type="checkbox" name="jewelry" value="plugs" id="filter-plugs" data-friendly="Plugs">
                    <label class="form-check-label d-block" for="filter-plugs">Plugs &amp; Tunnels
                    </label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input" type="checkbox" name="flare_type" value="Single flare" id="filter-single-flare" data-friendly="Single flare">
                    <label class="form-check-label d-block" for="filter-single-flare">Single flare</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input" type="checkbox" name="flare_type" value="Double flare" id="filter-double-flare" data-friendly="Double flare">
                    <label class="form-check-label d-block" for="filter-double-flare">Double flare</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input" type="checkbox" name="flare_type" value="No flare" id="filter-no-flare" data-friendly="No flares">
                    <label class="form-check-label d-block" for="filter-no-flare">No flare</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input" type="checkbox" name="flare_type" value="Screw on" id="filter-thread-on" data-friendly="Thread on back">
                    <label class="form-check-label d-block" for="filter-thread-on">Thread on back</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="saddle" id="filter-saddle" data-friendly="Saddle">
                    <label class="form-check-label d-block" for="filter-saddle">Saddles</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="customorders" value="customorder-yes" id="filter-preorderyes" data-friendly="Custom orders">
                    <label class="form-check-label d-block" for="filter-preorderyes">Custom orders</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="retainer" id="filter-retainers" data-friendly="Retainers">
                    <label class="form-check-label d-block" for="filter-retainers">Retainers</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="septum" id="filter-septum" data-friendly="Septum jewelry">
                    <label class="form-check-label d-block" for="filter-septum">Septum</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="barbell" id="filter-barbells" data-friendly="Straight barbells">
                    <label class="form-check-label d-block" for="filter-barbells">Straight Barbells</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="tapers" id="filter-tapers" data-friendly="Tapers">
                    <label class="form-check-label d-block" for="filter-tapers">Tapers / Stretching</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="twists" id="filter-twists" data-friendly="Twists">
                    <label class="form-check-label d-block" for="filter-twists">Twists</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="jewelry" value="weight" id="filter-weights" data-friendly="Weights">
                    <label class="form-check-label d-block" for="filter-weights">Weights</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input cat-select"  type="checkbox" name="jewelry" value="Hanging Designs" id="filter-hanging" data-name="hanging" data-friendly="Hanging Designs">
                    <label class="form-check-label d-block" for="filter-hanging">Hanging Designs (All)</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-hanging"  type="checkbox" name="jewelry" value="hoop" id="filter-hoops" data-friendly="Hoops">
                    <label class="form-check-label d-block" for="filter-hoops">Hoops</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-hanging"  type="checkbox" name="jewelry" value="ornate" id="filter-ornate" data-friendly="Ornate">
                    <label class="form-check-label d-block" for="filter-ornate">Ornate Shapes</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-hanging"  type="checkbox" name="jewelry" value="spiral" id="filter-spirals" data-friendly="Spirals">
                    <label class="form-check-label d-block" for="filter-spirals">Spirals</label>
                  </div>

                  <div class="h5 mt-3 mb-0 pb-1 w-75 border-bottom">Regular Jewelry</div>
                  <div class="form-check">
                    <input class="form-check-input cat-select"  type="checkbox" name="jewelry" value="Regular Jewelry" id="filter-regular" data-name="regjewelry" data-friendly="">
                    <label class="form-check-label d-block" for="filter-regular">Regular Jewelry (All)</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-regjewelry"  type="checkbox" name="jewelry" value="bracelet" id="filter-bracelets" data-friendly="Bracelets">
                    <label class="form-check-label d-block" for="filter-bracelets">Bracelets</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-regjewelry"  type="checkbox" name="jewelry" value="earring" id="filter-earring" data-friendly="Earrings">
                    <label class="form-check-label d-block" for="filter-earring">Earrings</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-regjewelry"  type="checkbox" name="jewelry" value="necklace" id="filter-necklace" data-friendly="Necklaces">
                    <label class="form-check-label d-block" for="filter-necklace">Necklaces</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-regjewelry"  type="checkbox" name="jewelry" value="chains-necklace" id="filter-necklace-chains" data-friendly="Necklace chains">
                    <label class="form-check-label d-block" for="filter-necklace-chains">Necklace chains</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-regjewelry"  type="checkbox" name="jewelry" value="finger-ring" id="filter-fingerrings" data-friendly="Finger rings">
                    <label class="form-check-label d-block" for="filter-fingerrings">Finger Rings</label>
                  </div>
                  <div class="h5 mt-3 mb-0 pb-1 w-75 border-bottom">
                    Other
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="jewelry" value="accessories" id="filter-accessories" data-friendly="Accessories">
                    <label class="form-check-label d-block" for="filter-accessories">Accessories</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="jewelry" value="aftercare" id="filter-aftercare" data-friendly="Aftercare products">
                    <label class="form-check-label d-block" for="filter-aftercare">Aftercare</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="jewelry" value="gear" id="filter-gear" data-friendly="Clothing &amp; Gear">
                    <label class="form-check-label d-block" for="filter-gear">Clothing / Gear</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="jewelry" value="stickers" id="filter-stickers" data-friendly="Stickers">
                    <label class="form-check-label d-block" for="filter-stickers">Stickers</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="jewelry" value="storage" id="filter-storage" data-friendly="Storage cases">
                    <label class="form-check-label d-block" for="filter-storage">Storage cases</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="jewelry" value="tools" id="filter-tools" data-friendly="Tools">
                    <label class="form-check-label d-block" for="filter-tools">Tools</label>
                  </div>
                </div>
              </div>
            </div>
           
          
            
            <div class="card rounded-0 border-0">
              <a class="card-header collapsed h6 filter-dropdown" id="gauge-head" data-toggle="collapse" data-target="#gauge" aria-expanded="false" aria-controls="gauge"
                href="#"><i class="fa" aria-hidden="true"></i>
                Gauges &amp; Sizes
              </a>
              <div id="gauge" class="collapse" aria-labelledby="gauge-head" data-parent="#accordion">
                <div class="card-body">
                 <%
' Get gauges to display in drop down menus
SqlString = "SELECT GaugeShow, conversion FROM TBL_GaugeOrder ORDER BY GaugeOrder ASC" 
Set rsGetGauges = DataConn.Execute(SqlString)
%>       

<% While NOT rsGetGauges.EOF %>

<% if Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) = "00g" then %>

		<div class="form-check">
			<input class="form-check-input cat-select" type="checkbox" name="gauge" value="00g" id="filter-gauge00g" data-name="00g" data-friendly="00g">
			 <label class="form-check-label d-block" for="filter-gauge00g">00g (all)</label>
		</div>
		<div class="form-check ml-4">
			<input class="form-check-input sub-00g" type="checkbox" name="gauge" value="00g/9mm" id="filter-00g-9mm" data-friendly="00g/9mm">
			 <label class="form-check-label d-block" for="filter-00g-9mm">00g/9mm</label>
		</div>
		<div class="form-check ml-4">
			<input class="form-check-input sub-00g" type="checkbox" name="gauge" value="00g/9.5mm" id="filter-00g-95mm" data-friendly="00g/9.5mm">
			 <label class="form-check-label d-block" for="filter-00g-95mm">00g/9.5mm</label>
		</div>
    		<div class="form-check ml-4">
			<input class="form-check-input sub-00g" type="checkbox" name="gauge" value="00g/10mm" id="filter-00g-10mm" data-friendly="00g/10mm">
			 <label class="form-check-label d-block" for="filter-00g-10mm">00g/10mm</label>
		</div>
<% end if %>

<% if Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) = "Youth large" then %>

		<div  class="h5 mt-3 mb-0 pb-1 w-75 border-bottom">Shirt sizes</div>

<% end if %>
<% if Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) = "Size 4" then %>
		<div  class="h5 mt-3 mb-0 pb-1 w-75 border-bottom">Ring sizes</div>

<% end if %>
<% if Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) <> "14g/12g" and Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) <> "18g/16g" and Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) <> "00g" and Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) <> "00g/9mm" and Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) <> "00g/9.5mm" and Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) <> "00g/10mm" and Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) <> " " and Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) <> "&nbsp;" and Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) <> Server.HTMLEncode(" ") and Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) <> "n/a" and Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) <> "" then %>
	<div class="form-check">
			<input class="form-check-input" type="checkbox" name="gauge" value="<%= Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) %>" id="filter-gauge<%= Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) %>" data-friendly="<%= Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) %>">
			 <label class="form-check-label d-block" for="filter-gauge<%= Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) %>"><%= Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) %><% if rsGetGauges.Fields.Item("conversion").Value <> "" then %> (<%= rsGetGauges.Fields.Item("conversion").Value %>)<% end if %></label>
		</div>
<% end if %>
<% rsGetGauges.MoveNext()
Wend 
rsGetGauges.ReQuery() %>	
                </div>
              </div>
            </div>
            <div class="card rounded-0 border-0">
              <a class="card-header collapsed h6 filter-dropdown" id="materials-head" data-toggle="collapse" data-target="#material" aria-expanded="false" aria-controls="material"
                href="#"><i class="fa" aria-hidden="true"></i>
                Material
              </a>
              <div id="material" class="collapse" aria-labelledby="material-head" data-parent="#accordion">
                <div class="card-body">
                 






		<div class="form-check h6">
			<input class="form-check-input cat-select" type="checkbox" name="exclude-material" value="on" id="filter-excludematerials" data-friendly="Exclude materials:">
			 <label class="form-check-label d-block" for="filter-excludematerials">Exclude materials:</label>
		</div>

	

		<div class="form-check">
			<input class="form-check-input cat-select" type="checkbox" name="material" value="Metals" id="filter-metals" data-name="metals" data-friendly="All metal jewelry">
      <label class="form-check-label d-block font-weight-bold" for="filter-metals">Metals (All)</label>
		</div>


	<div class="form-check ml-4">
		<input class="form-check-input sub-metals" type="checkbox" name="material" value="316L Stainless Steel" id="filter-allsteel" data-friendly="Steel">
    <label class="form-check-label d-block" for="filter-allsteel">Steel (Economy)</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-metals" type="checkbox" name="material" value="316LVM ASTM F-138 Implant Grade Steel" id="filter-impsteel"   data-friendly="Implant grade steel">
    <label class="form-check-label d-block" for="filter-impsteel">Steel (Implant grade)</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-metals" type="checkbox" name="material" value="Titanium" id="filter-alltitanium"   data-friendly="Titanium">
    <label class="form-check-label d-block" for="filter-alltitanium">Titanium (all)</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-metals" type="checkbox" name="material" value="Titanium implant grade" id="filter-imptitanium"   data-friendly="Implant grade titanium">
    <label class="form-check-label d-block" for="filter-imptitanium">Titanium (Implant grade)</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-metals" type="checkbox" name="material" value="Niobium" id="filter-niobium" data-friendly="Niobium">
    <label class="form-check-label d-block" for="filter-niobium">Niobium</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-metals" type="checkbox" name="material" value="Brass" id="filter-metbrass"  data-friendly="Brass">
    <label class="form-check-label d-block" for="filter-metbrass">Brass</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-metals" type="checkbox" name="material" value="Copper" id="filter-metcopper" data-friendly="Copper">
    <label class="form-check-label d-block" for="filter-metcopper">Copper</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-metals" type="checkbox" name="material" value="cobalt-chrome" id="filter-metcobaltchrome" data-friendly="Cobalt-Chrome">
    <label class="form-check-label d-block" for="filter-metcobaltchrome">Cobalt-Chrome</label>
	</div>


		<div class="form-check">
			<input class="form-check-input" type="checkbox" name="material" id="filter-glass" value="Glass" data-friendly="Glass">
      <label class="form-check-label d-block font-weight-bold" for="filter-glass">Glass</label>
		</div>



		<div class="form-check">
		<input class="form-check-input sub-plastic" type="checkbox" name="material" value="Silicone" id="filter-silicone" data-friendly="Silicone">
    <label class="form-check-label d-block font-weight-bold" for="filter-silicone">Silicone</label>
	</div>

	

		<div class="form-check">
			<input class="form-check-input cat-select" type="checkbox" name="material" value="Precious metals" data-name="precious" id="filter-preciousmetals" data-friendly="Precious metals">
      <label class="form-check-label d-block font-weight-bold" for="filter-preciousmetals">Precious metals (All)</label>
		</div>

	<div class="form-check ml-4">
		<input class="form-check-input sub-precious" type="checkbox" name="material" value="solid rose gold" id="filter-rosegold" data-friendly="Rose gold">
    <label class="form-check-label d-block" for="filter-rosegold">Rose Gold</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-precious" type="checkbox" name="material" value="solid white gold" id="filter-whitegold" data-friendly="White gold">
    <label class="form-check-label d-block" for="filter-whitegold">White Gold</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-precious" type="checkbox" name="material" value="solid yellow gold" id="filter-yellowgold" data-friendly="Yellow gold ml-4">
    <label class="form-check-label d-block" for="filter-yellowgold">Yellow Gold</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-precious" type="checkbox" name="material" value="Platinum" id="filter-platinum" data-friendly="Platinum">
    <label class="form-check-label d-block" for="filter-platinum">Platinum</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-precious" type="checkbox" name="material" value="Silver" id="filter-metsilver" data-friendly="Sterling silver">
    <label class="form-check-label d-block" for="filter-metsilver">Silver</label>
	</div>


		<div class="form-check">
			<input class="form-check-input cat-select" type="checkbox" name="material" value="Organics" id="filter-organics" data-name="organics" data-friendly="Organics">
      <label class="form-check-label d-block font-weight-bold" for="filter-organics">Organics (All)</label>
		</div>

	<div class="form-check ml-4">
		<input class="form-check-input sub-organics" type="checkbox" name="material" value="Amber" id="filter-orgamber" data-friendly="Amber">
    <label class="form-check-label d-block" for="filter-orgamber">Amber</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-organics" type="checkbox" name="material" value="Bone" id="filter-orgbone" data-friendly="Bone">
    <label class="form-check-label d-block" for="filter-orgbone">Bone</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-organics" type="checkbox" name="material" value="Horn" id="filter-horn" data-friendly="Horn">
    <label class="form-check-label d-block" for="filter-horn">Horn</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-organics" type="checkbox" name="material" value="Shell" id="filter-shell" data-friendly="Shell">
    <label class="form-check-label d-block" for="filter-shell">Shell</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-organics" type="checkbox" name="material" value="Stone" id="filter-stones" data-friendly="Stone">
    <label class="form-check-label d-block" for="filter-stones">Stone</label>
	</div>

		<div class="form-check">
			<input class="form-check-input cat-select sub-organics" type="checkbox" name="material" value="Wood" id="filter-woods" data-name="wood" data-friendly="Woods">
			<label class="form-check-label d-block font-weight-bold" for="filter-woods"> Wood (All types)</label>
		</div>
	
	<div class="form-check ml-4">
		<input class="form-check-input sub-wood sub-organics" type="checkbox" name="material" value="arang" id="filter-arang" data-friendly="Arang wood">
    <label class="form-check-label d-block" for="filter-arang">Arang</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-wood sub-organics" type="checkbox" name="material" value="bloodwood" id="filter-bloodwood" data-friendly="Bloodwood">
    <label class="form-check-label d-block" for="filter-bloodwood">Bloodwood</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-wood sub-organics" type="checkbox" name="material" value="crocodile" id="filter-croco" data-friendly="Crocodile wood">
    <label class="form-check-label d-block" for="filter-croco">Crocodile</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-wood sub-organics" type="checkbox" name="material" value="olivewood" id="filter-olivewood" data-friendly="Olivewood">
    <label class="form-check-label d-block" for="filter-olivewood">Olivewood</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-wood sub-organics" type="checkbox" name="material" value="saba" id="filter-saba" data-friendly="Saba wood">
    <label class="form-check-label d-block" for="filter-saba">Saba</label>
	</div>
	


		<div class="form-check">
			<input class="form-check-input cat-select" type="checkbox" name="cat-select" id="filter-matothers" data-name="other" data-friendly="">
      <label class="form-check-label d-block font-weight-bold" for="filter-matothers">Other</label>
		</div>
		
	<div class="form-check ml-4">
		<input class="form-check-input sub-other" type="checkbox" name="material" value="Acrylic" id="filter-acrylic" data-friendly="Acrylic">
    <label class="form-check-label d-block" for="filter-acrylic">Acrylic</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-other" type="checkbox" name="material" value="Bioplast" id="filter-bioplast" data-friendly="Bioplast">
    <label class="form-check-label d-block" for="filter-bioplast">Bioplast</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-other" type="checkbox" name="material" value="Flexible and plastic" id="filter-flexplastic" data-friendly="Flexible plastics">
    <label class="form-check-label d-block" for="filter-flexplastic">Flexible Plastic</label>
	</div>
	<div class="form-check ml-4">
		<input class="form-check-input sub-other" type="checkbox" name="material" value="PTFE" id="filter-ptfe" data-friendly="PTFE">
    <label class="form-check-label d-block" for="filter-ptfe">PTFE</label>
	</div>
                </div>
              </div>
            </div>
         
                <div class="card rounded-0 border-0">
              <a class="card-header collapsed h6 filter-dropdown" id="brand-head" data-toggle="collapse" data-target="#brand" aria-expanded="false" aria-controls="brand"
                href="#"><i class="fa" aria-hidden="true"></i>
                Brand
              </a>
              <div id="brand" class="collapse" aria-labelledby="brand-head" data-parent="#accordion">
                <div class="card-body">

<div class="form-check">
                    <input class="form-check-input cat-select" type="checkbox" name="brand" value="Premium Companies" data-name="premium" id="filter-premiumbrands" data-friendly="Premium metal companies">
                    <label class="form-check-label d-block" for="filter-premiumbrands">Premium Metals</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-premium" type="checkbox" name="brand" value="body circle" id="filter-bcd" data-friendly="Body Circle Designs">
                    <label class="form-check-label d-block" for="filter-bcd">Body Circle Designs</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-premium" type="checkbox" name="brand" value="element" id="filter-element" data-friendly="Element">
                    <label class="form-check-label d-block" for="filter-element">Element</label>
                  </div>
                  <div class="form-check ml-4">
                      <input class="form-check-input sub-premium" type="checkbox" name="brand" value="invictus" id="filter-invictus" data-friendly="Invictus">
                      <label class="form-check-label d-block" for="filter-invictus">Invictus</label>
                    </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-premium" type="checkbox" name="brand" value="le roi" id="filter-leroi" data-friendly="LeRoi">
                    <label class="form-check-label d-block" for="filter-leroi">Le Roi</label>
                  </div>
                   <div class="form-check ml-4">
                    <input class="form-check-input sub-premium" type="checkbox" name="brand" value="neometal" id="filter-neometal" data-friendly="Neometal">
                    <label class="form-check-label d-block" for="filter-neometal">Neometal</label>
                  </div>
                   <div class="form-check ml-4">
                    <input class="form-check-input sub-premium" type="checkbox" name="brand" value="sm316" id="filter-sm316" data-friendly="SM316">
                    <label class="form-check-label d-block" for="filter-sm316">SM316</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input cat-select" type="checkbox" name="brand" value="Economy Companies" data-name="economy" id="filter-economy" data-friendly="Economy brands">
                    <label class="form-check-label d-block" for="filter-economy">Economy</label>
                  </div>
                   <div class="form-check ml-4">
                    <input class="form-check-input sub-economy" type="checkbox" name="brand" value="body vibe" id="filter-bodyvibe" data-friendly="Body Vibe">
                    <label class="form-check-label d-block" for="filter-bodyvibe">Body Vibe</label>
                  </div>
                   <div class="form-check ml-4">
                    <input class="form-check-input sub-economy" type="checkbox" name="brand" value="half tone" id="filter-halftone" data-friendly="HalfTone">
                    <label class="form-check-label d-block" for="filter-halftone">Half Tone</label>
                  </div>
                   <div class="form-check ml-4">
                    <input class="form-check-input sub-economy" type="checkbox" name="brand" value="metal mafia" id="filter-metalmafia" data-friendly="Metal Mafia">
                    <label class="form-check-label d-block" for="filter-metalmafia">Metal Mafia</label>
                  </div>
                   <div class="form-check ml-4">
                    <input class="form-check-input sub-economy" type="checkbox" name="brand" value="wildcat" id="filter-wildcat" data-friendly="Wildcat">
                    <label class="form-check-label d-block" for="filter-wildcat">Wildcat</label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input cat-select" type="checkbox" name="brand" value="Glass Companies" data-name="glass" id="filter-glasscompanies" data-friendly="Glass companies">
                    <label class="form-check-label d-block" for="filter-glasscompanies">Glass</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-glass" type="checkbox" name="brand" value="atlas glass" id="filter-ag" data-friendly="Atlas Glass">
                    <label class="form-check-label d-block" for="filter-ag">Atlas Glass</label>
                  </div>
                   <div class="form-check ml-4">
                    <input class="form-check-input sub-glass" type="checkbox" name="brand" value="glasswear" id="filter-gws" data-friendly="Glasswear Studios">
                    <label class="form-check-label d-block" for="filter-gws">Glasswear Studios</label>
                  </div>
                   <div class="form-check ml-4">
                    <input class="form-check-input sub-glass" type="checkbox" name="brand" value="gorilla glass" id="filter-gg" data-friendly="Gorilla Glass">
                    <label class="form-check-label d-block" for="filter-gg">Gorilla Glass</label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="brand" value="kaos softwear" id="filter-kaos" data-friendly="Kaos Softwear">
                    <label class="form-check-label d-block" for="filter-kaos">Kaos Softwear</label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input cat-select" type="checkbox" name="brand" value="Gold Companies" data-name="goldbrands" id="filter-goldcompanies" data-friendly="Gold companies">
                    <label class="form-check-label d-block" for="filter-goldcompanies">Gold Brands</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-goldbrands" type="checkbox" name="brand" value="alchemy adornment" id="filter-alchemyadornment" data-friendly="Alchemy Adornment">
                    <label class="form-check-label d-block" for="filter-alchemyadornment">Alchemy Adornment</label>
                  </div>
                   <div class="form-check ml-4">
                    <input class="form-check-input sub-goldbrands" type="checkbox" name="brand" value="body gems" id="filter-bodygems" data-friendly="Body Gems">
                    <label class="form-check-label d-block" for="filter-bodygems">Body Gems</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-goldbrands two-filters" type="checkbox" name="brand" value="buddha jewelry" id="filter-buddha" data-filter2="material" data-filter2-value="Precious metals" data-friendly="Buddha Jewelry">
                    <label class="form-check-label d-block" for="filter-buddha">Buddha jewelry</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-goldbrands two-filters" type="checkbox" name="brand" value="invictus" id="filter-invictus" data-filter2="material" data-filter2-value="Precious metals" data-friendly="Invictus">
                    <label class="form-check-label d-block" for="filter-invictus">Invictus</label>
                  </div>
                   <div class="form-check ml-4">
                    <input class="form-check-input sub-goldbrands" type="checkbox" name="brand" value="venus by maria tash" id="filter-tash" data-friendly="Maria Tash">
                    <label class="form-check-label d-block" for="filter-tash">Maria Tash</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-goldbrands two-filters" data-filter2="material" data-filter2-value="Precious metals" type="checkbox" name="brand" value="maya organic" id="filter-maya2" data-friendly="Maya Jewelry">
                    <label class="form-check-label d-block" for="filter-maya2">Maya Jewelry</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-goldbrands two-filters" data-filter2="material" data-filter2-value="Precious metals" type="checkbox" name="brand" value="neometal" id="filter-neometal2" data-friendly="Neometal">
                    <label class="form-check-label d-block" for="filter-neometal2">Neometal</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-goldbrands two-filters" data-filter2="material" data-filter2-value="Precious metals" type="checkbox" name="brand" value="tawapa" id="filter-tawapa2" data-friendly="Tawapa">
                    <label class="form-check-label d-block" for="filter-tawapa2">Tawapa</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-goldbrands" type="checkbox" name="brand" value="norvoch" id="filter-alchemy" data-friendly="NorVoch">
                    <label class="form-check-label d-block" for="filter-norvoch">NorVoch</label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input cat-select" type="checkbox" name="brand" value="Organic &amp; Metals" data-name="organicbrands" id="filter-organicmetals" data-friendly="Organic &amp; Metal companies">
                    <label class="form-check-label d-block" for="filter-organicmetals">Organic &amp; Metals</label>
                  </div>
                   <div class="form-check ml-4">
                    <input class="form-check-input sub-organicbrands" type="checkbox" name="brand" value="buddha jewelry" id="filter-buddha" data-friendly="Buddha Jewelry">
                    <label class="form-check-label d-block" for="filter-buddha">Buddha jewelry</label>
                  </div>
                   <div class="form-check ml-4">
                    <input class="form-check-input sub-organicbrands" type="checkbox" name="brand" value="diablo organics" id="filter-diablo" data-friendly="Diablo Organics">
                    <label class="form-check-label d-block" for="filter-diablo">Diablo Organics</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-organicbrands" type="checkbox" name="brand" value="maya organic" id="filter-maya" data-friendly="Maya Jewelry">
                    <label class="form-check-label d-block" for="filter-maya">Maya Jewelry</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-organicbrands" type="checkbox" name="brand" value="oracle" id="filter-oracle" data-friendly="Oracle">
                    <label class="form-check-label d-block" for="filter-oracle">Oracle</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-organicbrands" type="checkbox" name="brand" value="quetzalli" id="filter-quetzalli" data-friendly="Quetzalli">
                    <label class="form-check-label d-block" for="filter-quetzalli">Quetzalli</label>
                  </div>
                  <div class="form-check ml-4">
                    <input class="form-check-input sub-organicbrands" type="checkbox" name="brand" value="tawapa" id="filter-tawapa" data-friendly="Tawapa">
                    <label class="form-check-label d-block" for="filter-tawapa">Tawapa</label>
                  </div>
                   <div class="form-check ml-4">
                    <input class="form-check-input sub-organicbrands" type="checkbox" name="brand" value="urban star" id="filter-urbanstar" data-friendly="Urban Star">
                    <label class="form-check-label d-block" for="filter-urbanstar">Urban Star</label>
                  </div>

                </div>
              </div>
            </div>

            <div class="card rounded-0 border-0">
              <a class="card-header collapsed h6 filter-dropdown" id="price-head" data-toggle="collapse" data-target="#price" aria-expanded="false" aria-controls="collapseOne"
                href="#collapseOne"><i class="fa" aria-hidden="true"></i>
                Price
              </a>
              <div id="price" class="collapse" aria-labelledby="price-head" data-parent="#accordion">
                <div class="card-body noscroll">

					<% 
					if(Instr(request.querystring("price"), ";")>0) then
						arrPrice = split(request.querystring("price"), ";")
					else
						arrPrice = Array(0, 100)
					end if	
					%>
						<input type="text" class="js-range-slider" name="price" id="price-range" value=""
						data-type="double"
						data-min="0"
						data-max="100"
						data-from="<%= arrPrice(0) %>"
						data-to="<%= arrPrice(1) %>"
						data-grid="true"
						/>

                </div>
              </div>
            </div>
                        <div class="card rounded-0 border-0">
              <a class="card-header collapsed h6 filter-dropdown" id="piercingtype-head" data-toggle="collapse" data-target="#piercingtype" aria-expanded="false" aria-controls="piercingtype"
                href="#"><i class="fa" aria-hidden="true"></i>
                Piercing Location
              </a>
              <div id="piercingtype" class="collapse" aria-labelledby="piercingtype-head" data-parent="#accordion">
                <div class="card-body">

                <div class="h5 mt-0 mb-0 pb-1 w-75 border-bottom">
                    Ear
                  </div>
                  <style>
                    #filters-ear-diagram #ear-diagram-image{width:50%}
                  </style>
                  <!--
                  <a class="btn btn-sm btn-purple text-light my-1" href="" data-toggle="modal" data-target="#modal-ear-diagram"
                  data-dismiss="modal" href="#">Click to enlarge diagram</a>
                  <div class="filters-ear-diagram">
                   virtual="/includes/ear-diagram-image.asp"
                  </div>-->
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Anti-tragus" id="filter-antitragus" data-friendly="Anti-tragus">
                    <label class="form-check-label d-block" for="filter-antitragus">Anti-tragus</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Basic ear piercing" id="filter-basicpiercing" data-friendly="Basic ear piercing">
                    <label class="form-check-label d-block" for="filter-basicpiercing">Basic ear piercing</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Conch" id="filter-conch" data-friendly="Conch">
                    <label class="form-check-label d-block" for="filter-conch">Conch</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Daith" id="filter-daith" data-friendly="Daith">
                    <label class="form-check-label d-block" for="filter-daith">Daith</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Helix" id="filter-helix" data-friendly="Helix">
                    <label class="form-check-label d-block" for="filter-helix">Helix</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Industrial" id="filter-industrial" data-friendly="Industrial">
                    <label class="form-check-label d-block" for="filter-industrial">Industrial</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Lobe" id="filter-lobe" data-friendly="Lobe">
                    <label class="form-check-label d-block" for="filter-lobe">Lobe</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Rook" id="filter-rook" data-friendly="Rook">
                    <label class="form-check-label d-block" for="filter-rook">Rook</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Snug" id="filter-snug" data-friendly="Snug">
                    <label class="form-check-label d-block" for="filter-snug">Snug</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Stretched lobe" id="filter-stretched" data-friendly="Stretched lobe">
                    <label class="form-check-label d-block" for="filter-stretched">Stretched lobe</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Tragus" id="filter-tragus" data-friendly="Tragus">
                    <label class="form-check-label d-block" for="filter-tragus">Tragus</label>
                  </div>
                  
<div class="h5 mt-3 mb-0 pb-1 w-75 border-bottom">
                    Nose
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Nostril" id="filter-nostril" data-friendly="Nostril">
                    <label class="form-check-label d-block" for="filter-nostril">Nostril</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Septum" id="filter-typeseptum" data-friendly="Septum">
                    <label class="form-check-label d-block" for="filter-typeseptum">Septum</label>
                  </div>
                  
<div class="h5 mt-3 mb-0 pb-1 w-75 border-bottom">
                    Other
                  </div>
<div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Microdermal" id="filter-microdermal" data-friendly="Microdermal">
                    <label class="form-check-label d-block" for="filter-microdermal">Microdermal</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Navel" id="filter-typenavel" data-friendly="Navel">
                    <label class="form-check-label d-block" for="filter-typenavel">Navel</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Nipple" id="filter-typenipple" data-friendly="Nipple">
                    <label class="form-check-label d-block" for="filter-typenipple">Nipple</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Surface" id="filter-surface" data-friendly="Surface">
                    <label class="form-check-label d-block" for="filter-surface">Surface</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Tongue" id="filter-tongue" data-friendly="Tongue">
                    <label class="form-check-label d-block" for="filter-tongue">Tongue</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Tongue web" id="filter-tongueweb" data-friendly="Tongue web">
                    <label class="form-check-label d-block" for="filter-tongueweb">Tongue web</label>
                  </div>
                  


                  
<div class="h5 mt-3 mb-0 pb-1 w-75 border-bottom">
                    Face
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Bites" id="filter-bites" data-friendly="Bites">
                    <label class="form-check-label d-block" for="filter-bites">Bites</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Bridge" id="filter-bridge" data-friendly="Bridge">
                    <label class="form-check-label d-block" for="filter-bridge">Bridge</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Cheek" id="filter-cheek" data-friendly="Cheek">
                    <label class="form-check-label d-block" for="filter-cheek">Cheek</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Eyebrow" id="filter-eyebrow" data-friendly="Eyebrow">
                    <label class="form-check-label d-block" for="filter-eyebrow">Eyebrow</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Jestrum" id="filter-jestrum" data-friendly="Jestrum">
                    <label class="form-check-label d-block" for="filter-jestrum">Jestrum</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Labret" id="filter-typelabret" data-friendly="Labret">
                    <label class="form-check-label d-block" for="filter-typelabret">Labret</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Lip" id="filter-typelip" data-friendly="Lip">
                    <label class="form-check-label d-block" for="filter-typelip">Lip</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Philtrum" id="filter-philtrum" data-friendly="Philtrum">
                    <label class="form-check-label d-block" for="filter-philtrum">Philtrum</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Vertical labret" id="filter-verticallabret" data-friendly="Vertical labret">
                    <label class="form-check-label d-block" for="filter-verticallabret">Vertical labret</label>
                  </div>
                  

                  
<div class="h5 mt-3 mb-0 pb-1 w-75 border-bottom">
                    Genital
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Ampallang" id="filter-ampallang" data-friendly="Ampallang">
                    <label class="form-check-label d-block" for="filter-ampallang">Ampallang</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Apadravya" id="filter-apadravya" data-friendly="Apadravya">
                    <label class="form-check-label d-block" for="filter-apadravya">Apadravya</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Clitoris" id="filter-clitoris" data-friendly="Clitoris">
                    <label class="form-check-label d-block" for="filter-clitoris">Clitoris</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Christina" id="filter-christina" data-friendly="Christina">
                    <label class="form-check-label d-block" for="filter-christina">Christina</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Dydoe" id="filter-dydoe" data-friendly="Dydoe">
                    <label class="form-check-label d-block" for="filter-dydoe">Dydoe</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Foreskin" id="filter-foreskin" data-friendly="Foreskin">
                    <label class="form-check-label d-block" for="filter-foreskin">Foreskin</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Fourchette" id="filter-fourchette" data-friendly="Fourchette">
                    <label class="form-check-label d-block" for="filter-fourchette">Fourchette</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Frenum" id="filter-frenum" data-friendly="Frenum">
                    <label class="form-check-label d-block" for="filter-frenum">Frenum</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Guiche" id="filter-guiche" data-friendly="Guiche">
                    <label class="form-check-label d-block" for="filter-guiche">Guiche</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Horizontal hood" id="filter-horhood" data-friendly="Horizontal hood">
                    <label class="form-check-label d-block" for="filter-horhood">Horizontal hood</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Labia" id="filter-labia" data-friendly="Labia">
                    <label class="form-check-label d-block" for="filter-labia">Labia</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Prince Albert" id="filter-pa" data-friendly="Prince Albert">
                    <label class="form-check-label d-block" for="filter-pa">Prince Albert</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Scrotum" id="filter-scrotum" data-friendly="Scrotum">
                    <label class="form-check-label d-block" for="filter-scrotum">Scrotum</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox" name="piercing" value="Vertical hood" id="filter-verhood" data-friendly="Vertical hood">
                    <label class="form-check-label d-block" for="filter-verhood">Vertical hood</label>
                  </div>


                </div>
              </div>
            </div>
                        <div class="card rounded-0 border-0">
              <a class="card-header collapsed h6 filter-dropdown" id="length-head" data-toggle="collapse" data-target="#length" aria-expanded="false" aria-controls="collapseOne"
                href="#collapseOne"><i class="fa" aria-hidden="true"></i>
                Length / Diameter
              </a>
              <div id="length" class="collapse" aria-labelledby="length-head" data-parent="#accordion">
                <div class="card-body">

          <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="4mm" id="filter-length4mm" data-friendly="Length: 4mm">
                    <label class="form-check-label d-block" for="filter-length4mm">4mm</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="3/16&quot;" id="filter-length3-16" data-friendly="Length: 3/16&quot;">
                    <label class="form-check-label d-block" for="filter-length3-16">3/16&quot; (4.7mm)</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="5mm" id="filter-length5mm" data-friendly="Length: 5mm">
                    <label class="form-check-label d-block" for="filter-length5mm">5mm</label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1/4&quot;" id="filter-length1-4" data-friendly="Length: 1/4&quot;">
                    <label class="form-check-label d-block" for="filter-length1-4">1/4&quot; (6.5mm)</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="7mm" id="filter-length7mm" data-friendly="Length: 7mm">
                    <label class="form-check-label d-block" for="filter-length7mm">7mm</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="9/32&quot;" id="filter-length9-32" data-friendly="Length: 9/32&quot;">
                    <label class="form-check-label d-block" for="filter-length9-32">9/32&quot; </label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="5/16&quot;" id="filter-length5-16" data-friendly="Length: 5/16&quot;">
                    <label class="form-check-label d-block" for="filter-length5-16">5/16&quot; (8mm)</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="11/32&quot;" id="filter-length11-32" data-friendly="Length: 11/32&quot; (9mm)">
                    <label class="form-check-label d-block" for="filter-length11-32">11/32&quot; (9mm)</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="3/8&quot;" id="filter-length3-8" data-friendly="Length: 3/8&quot; (9.5mm)">
                    <label class="form-check-label d-block" for="filter-length3-8">3/8&quot; (9.5mm)</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="10mm" id="filter-length10mm" data-friendly="Length: 10mm">
                    <label class="form-check-label d-block" for="filter-length10mm">10mm</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="13/32&quot;" id="filter-length13-32" data-friendly="Length: 13/32&quot;">
                    <label class="form-check-label d-block" for="filter-length13-32">13/32&quot; </label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="7/16&quot;" id="filter-length7-16" data-friendly="Length: 7/16&quot;">
                    <label class="form-check-label d-block" for="filter-length7-16">7/16&quot; (~11mm)</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="11mm" id="filter-length11mm" data-friendly="Length: 11mm">
                    <label class="form-check-label d-block" for="filter-length11mm">11mm</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="12mm" id="filter-length12mm" data-friendly="Length: 12mm">
                    <label class="form-check-label d-block" for="filter-length12mm">12mm</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1/2&quot;" id="filter-length1-2" data-friendly="Length: 1/2&quot;">
                    <label class="form-check-label d-block" for="filter-length1-2">1/2&quot;  (12.5mm)</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="9/16&quot;" id="filter-length9-16" data-friendly="Length: 9/16&quot;">
                    <label class="form-check-label d-block" for="filter-length9-16">9/16&quot;  (14mm)</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="5/8&quot;" id="filter-length5-8" data-friendly="Length: 5/8&quot;">
                    <label class="form-check-label d-block" for="filter-length5-8">5/8&quot;  (16mm)</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="11/16&quot;" id="filter-length11-16" data-friendly="Length: 11/16&quot;">
                    <label class="form-check-label d-block" for="filter-length11-16">11/16&quot;  (~18mm)</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="3/4&quot;" id="filter-length3-4" data-friendly="Length: 3/4&quot;">
                    <label class="form-check-label d-block" for="filter-length3-4">3/4&quot;  (19mm)</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="13/16&quot;" id="filter-length13-16" data-friendly="Length: 13/16&quot;">
                    <label class="form-check-label d-block" for="filter-length13-16">13/16&quot;  (20.5mm)</label>
                  </div>
                  <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="7/8&quot;" id="filter-length7-8" data-friendly="Length: 7/8&quot;">
                    <label class="form-check-label d-block" for="filter-length7-8">7/8&quot;  (22mm)</label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="15/16&quot;" id="filter-length15-16" data-friendly="Length: 15/16&quot;">
                    <label class="form-check-label d-block" for="filter-length15-16">15/16&quot;  (24mm)</label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1&quot;" id="filter-length1inch" data-friendly="Length: 1&quot;">
                    <label class="form-check-label d-block" for="filter-length1inch">1&quot;  (~25mm)</label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1-1/8&quot;" id="filter-length1-18" data-friendly="Length: 1-1/8&quot;">
                    <label class="form-check-label d-block" for="filter-length1-18">1-1/8&quot; (28.5mm)</label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1-3/16&quot;" id="filter-length1-316" data-friendly="Length: 1-3/16&quot;">
                    <label class="form-check-label d-block" for="filter-length1-316">1-3/16&quot;</label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1-1/4&quot;" id="filter-length1-14" data-friendly="Length: 1-1/4&quot;">
                    <label class="form-check-label d-block" for="filter-length1-14">1-1/4&quot; (32mm)</label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1-5/16&quot;" id="filter-length1-516" data-friendly="Length: 1-5/16&quot;">
                    <label class="form-check-label d-block" for="filter-length1-516">1-5/16&quot; (33mm)</label>
                  </div>
                     <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1-3/8&quot;" id="filter-length1-38" data-friendly="Length: 1-3/8&quot;">
                    <label class="form-check-label d-block" for="filter-length1-38">1-3/8&quot; (35mm)</label>
                  </div>
                     <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1-7/16&quot;" id="filter-length1-716" data-friendly="Length: 1-7/16&quot;">
                    <label class="form-check-label d-block" for="filter-length1-716">1-7/16&quot; (36.5mm)</label>
                  </div>
                     <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1-1/2&quot;" id="filter-length1-12" data-friendly="Length: 1-1/2&quot;">
                    <label class="form-check-label d-block" for="filter-length1-12">1-1/2&quot; (38mm)</label>
                  </div>
                     <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1-9/16&quot;" id="filter-length1-916" data-friendly="Length: 1-9/16&quot;">
                    <label class="form-check-label d-block" for="filter-length1-916">1-9/16&quot; (40mm)</label>
                  </div>
                     <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1-5/8&quot;" id="filter-length1-58" data-friendly="Length: 1-5/8&quot;">
                    <label class="form-check-label d-block" for="filter-length1-58">1-5/8&quot; (41mm)</label>
                  </div>
                     <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1-3/4&quot;" id="filter-length1-34" data-friendly="Length: 1-3/4&quot;">
                    <label class="form-check-label d-block" for="filter-length1-34">1-3/4&quot; (44mm)</label>
                  </div>
                     <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="1-7/8&quot;" id="filter-length1-78" data-friendly="Length: 1-7/8&quot;">
                    <label class="form-check-label d-block" for="filter-length1-78">1-7/8&quot; (48mm)</label>
                  </div>
                     <div class="form-check">
                    <input class="form-check-input" type="checkbox" name="length" value="2&quot;" id="filter-length2inch" data-friendly="Length: 2&quot;">
                    <label class="form-check-label d-block" for="filter-length2inch">2&quot; (51mm)</label>
                  </div>

                </div>
              </div>
            </div>
                        <div class="card rounded-0 border-0">
              <a class="card-header collapsed h6 filter-dropdown" id="sales-head" data-toggle="collapse" data-target="#sales" aria-expanded="false" aria-controls="collapseOne"
                href="#collapseOne"><i class="fa" aria-hidden="true"></i>
                Sales - Save $$
              </a>
              <div id="sales" class="collapse" aria-labelledby="sales-head" data-parent="#accordion">
                <div class="card-body noscroll">


	<div class="form-check">
		<input class="form-check-input" type="radio" name="discount" value="all" id="filter-allsales" data-friendly="All sales">
    <label class="form-check-label d-block" for="filter-allsales">All sales</label>
	</div>
  <div class="form-check">
		<input class="form-check-input" type="radio" name="discount" value="5-20" id="filter-sale10-20" data-friendly="10% - 20% off">
    <label class="form-check-label d-block" for="filter-sale10-20">10% - 20% off</label>
	</div>
  <div class="form-check">
		<input class="form-check-input" type="radio" name="discount" value="25-45" id="filter-sale25-45" data-friendly="25% - 45% off">
    <label class="form-check-label d-block" for="filter-sale25-45">25% - 45% off</label>
	</div>
  <div class="form-check">
		<input class="form-check-input" type="radio" name="discount" value="50-70" id="filter-sale50-70" data-friendly="50% - 70% off">
    <label class="form-check-label d-block" for="filter-sale50-70">50% - 70% off</label>
	</div>
  <div class="form-check">
		<input class="form-check-input" type="radio" name="discount" value="75-90" id="filter-sale75up" data-friendly="75% + off">
    <label class="form-check-label d-block" for="filter-sale75up">75% + off</label>
	</div>

                </div>
              </div>
            </div>
                        <div class="card rounded-0 border-0">
              <a class="card-header collapsed h6 filter-dropdown" id="threading-head" data-toggle="collapse" data-target="#threading" aria-expanded="false" aria-controls="collapseOne"
                href="#collapseOne"><i class="fa" aria-hidden="true"></i>
                Threading
              </a>
              <div id="threading" class="collapse" aria-labelledby="threading-head" data-parent="#accordion">
                <div class="card-body noscroll">

                 <div class="form-check">
                    <input class="form-check-input"  type="checkbox"  name="threading" value="Externally threaded" id="filter-external" data-friendly="Externally threaded">
                    <label class="form-check-label d-block" for="filter-external">Externally threaded</label>
                  </div>
                                   <div class="form-check">
                    <input class="form-check-input"  type="checkbox"  name="threading" value="Internally threaded" id="filter-internal" data-friendly="Internally threaded">
                    <label class="form-check-label d-block" for="filter-internal">Internally threaded</label>
                  </div>
                                   <div class="form-check">
                    <input class="form-check-input"  type="checkbox"  name="threading" value="Threadless" id="filter-threadless" data-friendly="Threadless">
                    <label class="form-check-label d-block" for="filter-threadless">Threadless</label>
                  </div>


                </div>
              </div>
            </div>
                        <div class="card rounded-0 border-0">
              <a class="card-header collapsed h6 filter-dropdown" id="customorders-head" data-toggle="collapse" data-target="#customorders" aria-expanded="false" aria-controls="collapseOne"
                href="#collapseOne"><i class="fa" aria-hidden="true"></i>
                Custom Orders
              </a>
              <div id="customorders" class="collapse" aria-labelledby="customorders-head" data-parent="#accordion">
                <div class="card-body noscroll">

 <div class="form-check">
                    <input class="form-check-input"   type="radio" name="customorders" value="customorder-not" id="filter-nocustomorders" data-friendly="Do not show custom orders">
                    <label class="form-check-label d-block" for="filter-nocustomorders">Do not show custom orders</label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input" type="radio" name="customorders" value="customorder-yes" id="filter-onlycustomorders" data-friendly="Only show custom orders">
                    <label class="form-check-label d-block" for="filter-onlycustomorders">Show ONLY custom orders</label>
                  </div>

                </div>
              </div>
            </div>
                        <div class="card rounded-0 border-0">
              <a class="card-header collapsed h6 filter-dropdown" id="singles-head" data-toggle="collapse" data-target="#singles" aria-expanded="false" aria-controls="collapseOne"
                href="#collapseOne"><i class="fa" aria-hidden="true"></i>
                Singles / Pairs
              </a>
              <div id="singles" class="collapse" aria-labelledby="singles-head" data-parent="#accordion">
                <div class="card-body noscroll">

 <div class="form-check">
                    <input class="form-check-input"   type="radio" name="pair" value="pairs" id="filter-onlypairs" data-friendly="Only show pairs">
                    <label class="form-check-label d-block" for="filter-onlypairs">Only show pairs</label>
                  </div>
                   <div class="form-check">
                    <input class="form-check-input" type="radio" name="pair" value="singles" id="filter-onlysingles" data-friendly="Only show singles">
                    <label class="form-check-label d-block" for="filter-onlysingles">Only show singles</label>
                  </div>


                </div>
              </div>
            </div>
                        <div class="card rounded-0 border-0">
              <a class="card-header collapsed h6 filter-dropdown" id="morefilters-head" data-toggle="collapse" data-target="#morefilters" aria-expanded="false" aria-controls="collapseOne"
                href="#collapseOne"><i class="fa" aria-hidden="true"></i>
                More Filters
              </a>
              <div id="morefilters" class="collapse" aria-labelledby="morefilters-head" data-parent="#accordion">
                <div class="card-body noscroll">
                                  <div class="form-check">
                    <input class="form-check-input"  type="checkbox"  name="limited" value="yes" id="filter-limitedonly" data-friendly="Limited items only">
                    <label class="form-check-label d-block" for="filter-limitedonly">Limited items only</label>
                  </div>
                                    <div class="form-check">
                    <input class="form-check-input"  type="checkbox"  name="onetime" value="yes" id="filter-onlyoneoff" data-friendly="One off items only">
                    <label class="form-check-label d-block" for="filter-onlyoneoff">Show only one-offs</label>
                  </div>
                </div>
              </div>
            </div>
            <%
'Create array for drop down menus
color_array = array("amber", "aqua", "black", "blue", "bone", "brass", "bronze", "brown", "clear", "copper", "dark-blue", "dark-purple", "fuchsia", "hider", "image", "iridescent", "gold", "glow", "gray", "green", "lavender", "light-blue", "lime", "magenta", "metallic", "navy", "neon", "opalescent", "orange", "pattern", "pink", "purple", "rainbow", "red", "rose-gold", "silver", "skin-tone", "tan", "teal", "translucent", "turquoise", "white", "yellow")
%>
                        <div class="card rounded-0 border-0">
              <a class="card-header collapsed h6 filter-dropdown" id="color-head" data-toggle="collapse" data-target="#color" aria-expanded="false" aria-controls="collapseOne"
                href="#collapseOne"><i class="fa" aria-hidden="true"></i>
                Color
              </a>
              <div id="color" class="collapse" aria-labelledby="color-head" data-parent="#accordion">
                <div class="card-body">
                    <div class="small text-secondary">
                      Defaults to contain ANY selected colors
                    </div>
                    <div class="form-check mb-3">
                      <input class="form-check-input" type="checkbox" name="color-filter" value="and" id="filter-color-and" data-friendly="Contains ALL selected colors">
                      <label class="form-check-label d-block small text-secondary" for="filter-color-and">Contains ALL selected colors</label>
                    </div>
                 <% for each x in color_array %>
  <div class="form-check">
		<input class="form-check-input" type="checkbox" name="colors" value="<%= x %>" id="filter-<%= x %>" data-friendly="<%= x %>">
    <label class="form-check-label d-block" for="filter-<%= x %>"><%= x %></label>
	</div>
	<% next %>
                </div>
              </div>
            </div>
            </div>
            <div class="p-3 w-100 m-0 bg-dark position-search-bottom">
            <button class="btn btn-purple w-100" type="submit">SEARCH &AMP; APPLY FILTERS</button>
            <span id="filter-builder-text" class="text-light small"></span>
            </div>
        </form>
      </div>
      <div class="col-12 col-lg-9 col-xl-10 pt-4 pl-lg-4 pr-lg-3 px-2" id="body-column">

        <%
        if request.cookies("OrderAddonsActive") <> "" then 
        
        ' ===========================================================================
        ' Check status of order to make sure it's still eligible to add items to it
        ' ===========================================================================

        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "SELECT ID, shipped, ship_code, ScanInvoice_Timestamp FROM sent_items WHERE ID = ?"
        objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,10, request.cookies("OrderAddonsActive")))
        Set rsVerifyAddonOrder = objCmd.Execute()
        
        if rsVerifyAddonOrder.Fields.Item("ship_code").Value = "paid" AND ISNULL(rsVerifyAddonOrder.Fields.Item("ScanInvoice_Timestamp").Value) AND (rsVerifyAddonOrder.Fields.Item("shipped").Value = "CUSTOM ORDER IN REVIEW" OR rsVerifyAddonOrder.Fields.Item("shipped").Value = "Pending..." OR rsVerifyAddonOrder.Fields.Item("shipped").Value = "Review" OR rsVerifyAddonOrder.Fields.Item("shipped").Value = "Pending shipment") then
        %>
        <div class="alert alert-info" id="addon-alert">
          <h4>ADDING ITEMS TO YOUR ORDER #<%= request.cookies("OrderAddonsActive") %></h4>
          You are currently adding items to an order that has not shipped yet. If you no longer want to add items OR you want to place a separate order, click the button below.<br/> <button class="btn btn-info mt-2" id="btn-cancel-addons">CLICK HERE TO CANCEL ADDING ITEMS</button>
        </div>
        <% 
        else 
          response.cookies("OrderAddonsActive") = ""
          Response.Cookies("OrderAddonsActive").Expires = DATE-1
          %>
          <div class="alert alert-danger">
            <h4>ORDER NOT ELIGIBLE FOR ADD-ONS</h4>
            This order is currently in <strong><%= rsVerifyAddonOrder.Fields.Item("shipped").Value %></strong> state and can not have items added to it. Only orders that have not shipped can have items added.<br><br>If you need to speak with someone feel free to <a href="/contact.asp">contact us</a> :)
          </div>
          <%
          end if ' on the correct order status
        end if ' addon cookie not null %>

        <!-- Modal for ear diagram -->
        <div class="modal fade filters-ear-diagram" id="modal-ear-diagram" tabindex="-1" role="dialog" aria-labelledby="LabelEarDiagram"
                aria-hidden="true">
                <div class="modal-dialog" role="document">
                        <div class="modal-content">
                                <div class="modal-header">
                                        <div class="modal-title" id="LabelEarDiagram">
                                          <h5>Common ear piercing locations</h5>
                                          Select a tag to apply it to your filter
                                        </div>
                                        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                                                <span aria-hidden="true">&times;</span>
                                        </button>
                                </div>
                    
                               <!--#include virtual="/includes/ear-diagram-image.asp"-->
                               <div class="modal-footer">
                                <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
                              </div>
                        </div>
                </div>
        </div>