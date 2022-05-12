<!--#include virtual="/functions/security.inc" -->
<%
' THIS IS THE NEW MOBILE SEARCH FILE

' ORDER BY DROP DOWN SELECT MENU

	if Request.form("sortby") <> "" then
		Session("filter_orderby") = Request.form("sortby")
		OrderSplit = Split(Session("filter_orderby"))
		OrderBy = OrderSplit(0)
		DirectionOrder = OrderSplit(1)
	else
		Session("filter_orderby") = Session("filter_orderby")
			If Session("filter_orderby") <> "" then
				OrderSplit = Split(Session("filter_orderby"))
				OrderBy = OrderSplit(0)
				DirectionOrder = OrderSplit(1)
			else
			   If request.form("brand") = "" then
				OrderBy = "new_page_date" 'min_price
				DirectionOrder = "desc" 'asc
			   else
				OrderBy = "new_page_date"
				DirectionOrder = "desc"			   	
			   end if
			end if
	end if
	
	If request.querystring("restock") = "restock" then
		OrderBy = "daterestocked"
		DirectionOrder = "desc"	
	end if	





' SET VARIABLES
DetectTags = ""

' Create array for category selection
if request.querystring("jewelry") <> "" then

	category_cleaned = request.querystring("jewelry")
	category_cleaned = replace(category_cleaned, "(", "")
	category_cleaned = replace(category_cleaned, ")", "")
	category_cleaned = replace(category_cleaned, "[", "")
	category_cleaned = replace(category_cleaned, "]", "")
	category_cleaned = replace(category_cleaned, """", "")
	category_cleaned = replace(category_cleaned, "rings", "finger-ring")
	category_cleaned = replace(category_cleaned, "ofinger-ring", "orings") ' messes up orings so fix back

	'If more than one checkbox selected
	if Instr(category_cleaned, ",") > 0 then
		
		category_array = split(replace(category_cleaned, "-", ""), ",")
		var_build_category = ""
		For i = 0 to Ubound(category_array)
			if i = 0 then
				var_build_category = var_build_category + "jewelry:" + replace(category_array(i), " ", "+")
			else
				var_build_category = var_build_category + " or jewelry:" + replace(category_array(i), " ", "+")
			end if
		next
		var_category = var_build_category
	else ' if only one category selected
		var_category = "jewelry:" + replace(replace(category_cleaned, "-", ""), " ", "+")	
	end if
	
	if Instr(category_cleaned, "Hanging Designs") > 0  then
			var_category = var_category & " or jewelry:hoop or jewelry:ornate or jewelry:pincher or jewelry:spiral or jewelry:plugloops"
	end if 
	
	if Instr(category_cleaned, "Regular Jewelry") > 0  then
			var_category = var_category & " or jewelry:bracelet or jewelry:earring or jewelry:necklace or jewelry:finger-rings"
	end if 
	if category_cleaned = "earring" then
			var_category = var_category & " or jewelry:earring or jewelry:earringstud or jewelry:earringdangle or jewelry:earringhuggies or jewelry:earcuff"
	end if 
	
	if Instr(category_cleaned, "navel") > 0  then
			var_category = var_category & " or jewelry:belly"
	end if 

	if category_cleaned = "captive" then
			var_category = var_category & " or jewelry:captive or jewelry:clicker"
	end if 
	
	DetectTags = "yes"
	
	' Special situation that if category is basics then it needs to be AND and not OR
	var_category = replace(var_category, "basics or", "basics and")
	var_category = replace(var_category, "or jewelry:+basics", "and jewelry:+basics")
	var_category = replace(var_category, "jewelry:customorderyes", "customorder-yes")
	var_category = replace(var_category, "jewelry:+customorderyes", "customorder-yes")
	
	var_category = "(" + var_category + ")"
else ' if no category was selected
	var_category = ""
end if ' array for category selection


' Create array for GAUGE selection
if request.querystring("gauge") <> "" then

	gauge_cleaned = request.querystring("gauge")
	gauge_cleaned = replace(gauge_cleaned, "(", "")
	gauge_cleaned = replace(gauge_cleaned, ")", "")
	gauge_cleaned = replace(gauge_cleaned, "[", "")
	gauge_cleaned = replace(gauge_cleaned, "]", "")


	'If more than one checkbox selected
	if Instr(gauge_cleaned, ",") > 0 then
		
		gauge_array = split(gauge_cleaned, ",")
		var_build_gauge = ""
		For i = 0 to Ubound(gauge_array)
			if i = 0 then
				var_build_gauge = var_build_gauge + "gauge:" + replace(gauge_array(i), " ", "")
			else
				var_build_gauge = var_build_gauge + " or gauge:" + replace(gauge_array(i), " ", "")
			end if
		next
		var_gauge = var_build_gauge
	else ' if only one gauge selected
		var_gauge = "gauge:" + replace(gauge_cleaned, " ", "")
	end if
	
	if Instr(gauge_cleaned, "14g") > 0 or Instr(gauge_cleaned, "12g") > 0 then
			var_gauge = var_gauge & " or gauge:14g/12g"
	end if
	if Instr(gauge_cleaned, "18g") > 0 or Instr(gauge_cleaned, "16g") > 0  then
			var_gauge = var_gauge & " or gauge:18g/16g"
	end if
	
	var_gauge = "(" & var_gauge & ")"
	
	if detect_detail_tags = "yes" then
		var_gauge = " and " + var_gauge
	end if
	
	var_gauge = replace(replace(replace(var_gauge, "/", "s"), """", "i"), "-", "h")
	detect_detail_tags = "yes"
else ' if no category was selected
	var_gauge = ""
end if ' array for GAUGE selection

' Create array for MATERIAL selection
if request.querystring("material") <> ""  then

	material_cleaned = request.querystring("material")
	material_cleaned = replace(material_cleaned, "(", "")
	material_cleaned = replace(material_cleaned, ")", "")
	material_cleaned = replace(material_cleaned, "[", "")
	material_cleaned = replace(material_cleaned, "]", "")
	material_cleaned = replace(material_cleaned, """", "")

	'If more than one checkbox selected
	if Instr(material_cleaned, ",") > 0 then
		
		material_array = split(material_cleaned, ",")
		var_build_material = ""
		For i = 0 to Ubound(material_array)
			if i = 0 then
				var_build_material = var_build_material +  replace(material_array(i), " ", "+")
			else
				var_build_material = var_build_material + " or " + replace(material_array(i), " ", "+")
			end if
		next
		var_material = var_build_material
	else ' if only one category selected
		var_material = replace(material_cleaned, " ", "+")
	end if

	if Instr(material_cleaned, "316L Stainless Steel") > 0  then
			var_material = var_material & " or Color+plated+steel or PVD+plated+steel"
	end if

	if Instr(material_cleaned, "Titanium") > 0 AND Instr(material_cleaned, "Titanium implant grade") < 0  then
			var_material = var_material & " or Color+plated+titanium or PVD+plated+titanium"
	end if

	if Instr(material_cleaned, "Organics") > 0  then
			var_material = var_material & " or Amber or Bamboo or Bone or Fossils or Fossilized+bone or Horn or Jet or Palm+seed or Shell or Stone or vegan or wood or Amboyna+burl or Apricot or Arang or Bloodwood or Chechen or Cherry or Coconut+wood or Crocodile or Ebony or Gentawas or Granadillo or Ironwood or Jackfruit or Katalox or Lignum+vitae or Maple or Olivewood or Osage+orange or Palmwood or Pear or Pink+ivory or Saba or Tamarind or Tewel or Tiger+ebony or Verawood or Zebrawood"
	end if 

	if Instr(material_cleaned, "Wood") > 0  then
			var_material = var_material & " or wood or Amboyna+burl or Apricot or Arang or Bloodwood or Chechen or Cherry or Coconut+wood or Crocodile or Ebony or Gentawas or Granadillo or Ironwood or Jackfruit or Katalox or Lignum+vitae or Maple or Olivewood or Osage+orange or Palmwood or Pear or Pink+ivory or Saba or Tamarind or Tewel or Tiger+ebony or Verawood or Zebrawood"
	end if

	if Instr(material_cleaned, "Metals") > 0  then
			var_material = var_material & " or steel or titanium or niobium or copper or brass or gold or silver or platinum or bronze or metal"
	end if 	
	
	if Instr(material_cleaned, "Precious metals") > 0  then
			var_material = var_material & " or solid+rose+gold or solid+white+gold or solid+yellow+gold or platinum"
	end if 	
	
	var_material = "(" & var_material & ")"
	
else ' if no category was selected
	var_material = ""
end if ' array for MATERIALS selection

	' Alter material for implant grade titanium to search BOTH implant grade options that we have
	var_material = replace(var_material, "Titanium+implant+grade", "Titanium+6AL4VELI+ASTM+F-136+implant+grade or Titanium+Ti-6Al-7Nb+ASTM-F1295+implant+grade or Titanium+ASTM+F136")
	var_material = replace(var_material, "(", "")
	var_material = replace(var_material, ")", "")
	var_material = replace(var_material, "[", "")
	var_material = replace(var_material, "]", "")

	


' Create array for BRAND selection
if request.querystring("brand") <> "" then

	brand_cleaned = request.querystring("brand")
	brand_cleaned = replace(brand_cleaned, "(", "")
	brand_cleaned = replace(brand_cleaned, ")", "")
	brand_cleaned = replace(brand_cleaned, "[", "")
	brand_cleaned = replace(brand_cleaned, "]", "")
	brand_cleaned = replace(brand_cleaned, """", "")

	'If more than one checkbox selected
	if Instr(brand_cleaned, ",") > 0 then
		
		brand_array = split(brand_cleaned, ",")
		var_build_brand = ""
		For i = 0 to Ubound(brand_array)
			if i = 0 then
				var_build_brand = var_build_brand +  replace(brand_array(i), " ", "+")
			else
				var_build_brand = var_build_brand + " or " + replace(brand_array(i), " ", "+")
			end if
		next
		
		var_brand = var_build_brand
	else ' if only one category selected
		var_brand = replace(brand_cleaned, " ", "+")
	end if
	
	if Instr(brand_cleaned, "Premium Companies") > 0  then
			var_brand = var_brand & " or Anatometal or Body+Circle+Designs or Industrial+Strength or Intrinsic or Le+Roi or Neometal or SM316"
	end if 		
	
	if Instr(brand_cleaned, "Economy Companies") > 0  then
			var_brand = var_brand & " or Body+Vibe or Half+Tone or Metal+Mafia or Wildcat"
	end if 	
	
	if Instr(brand_cleaned, "Glass Companies") > 0  then
			var_brand = var_brand & " or Glass+heart+studio or Glasswear+Studios or Gorilla+Glass or Modifika"
	end if 
	
	if Instr(brand_cleaned, "Glass Companies") > 0  then
			var_brand = var_brand & " or Glass+heart+studio or Glasswear+Studios or Gorilla+Glass or Modifika"
	end if 

	if Instr(brand_cleaned, "Organic & Metals") > 0  then
			var_brand = var_brand & " or Buddha+jewelry or Diablo+organics or Little7 or Maya+organic or Omerica or Oracle or Quetzalli or Tawapa or Tether or Urban+Star"
	end if 	

	if Instr(brand_cleaned, "Gold Companies") > 0  then
			var_brand = var_brand & " or Body+Gems or Venus+by+Maria+Tash  or solid+rose+gold or solid+white+gold or solid+yellow+gold or platinum"
	end if 	

	var_brand = "(" + var_brand + ")"
	
	if DetectTags = "yes" then
		var_brand = " and " + var_brand
	end if
	
	DetectTags = "yes"
else ' if no category was selected
	var_brand = ""
end if ' array for BRAND selection


' Create array for LENGTH selection
if request.querystring("length") <> "" then

	length_cleaned = request.querystring("length")
	length_cleaned = replace(length_cleaned, "(", "")
	length_cleaned = replace(length_cleaned, ")", "")
	length_cleaned = replace(length_cleaned, "[", "")
	length_cleaned = replace(length_cleaned, "]", "")
	
	'If more than one checkbox selected
	if Instr(length_cleaned, ",") > 0 then
		
		length_array = split(length_cleaned, ",")
		var_build_length = ""
		For i = 0 to Ubound(length_array)
			if i = 0 then
				var_build_length = var_build_length + " length:" + replace(length_array(i), " ", "+")
			else
				var_build_length = var_build_length + " or length:" + replace(length_array(i), " ", "+")
			end if
		next
		var_length = "(" + var_build_length + ")"
	else ' if only one category selected
		var_length = "(length:" + replace(length_cleaned, " ", "+") + ")"
	end if
	
	if detect_detail_tags = "yes" then
		var_length = " and " + var_length
	end if
	
	var_length = replace(replace(replace(var_length, "/", "s"), """", "i"), "-", "h")
	detect_detail_tags = "yes"
else ' if no category was selected
	var_length = ""
end if ' array for LENGTH selection

' Create array for PIERCING TYPE selection
if request.querystring("piercing") <> "" then

	piercing_cleaned = request.querystring("piercing")
	piercing_cleaned = replace(piercing_cleaned, "(", "")
	piercing_cleaned = replace(piercing_cleaned, ")", "")
	piercing_cleaned = replace(piercing_cleaned, "[", "")
	piercing_cleaned = replace(piercing_cleaned, "]", "")
	piercing_cleaned = replace(piercing_cleaned, """", "")

	'If more than one checkbox selected
	if Instr(piercing_cleaned, ",") > 0 then
		
		piercing_array = split(piercing_cleaned, ",")
		var_build_piercing = ""
		For i = 0 to Ubound(piercing_array)
			if i = 0 then
				var_build_piercing = var_build_piercing +  replace(piercing_array(i), " ", "+")
			else
				var_build_piercing = var_build_piercing + " or " + replace(piercing_array(i), " ", "+")
			end if
		next
		var_piercing = "(" + var_build_piercing + ")"
	else ' if only one category selected
		var_piercing = "(" + replace(piercing_cleaned, " ", "+") + ")"
	end if
	
	if DetectTags = "yes" then
		var_piercing = " and " + var_piercing
	end if
	
	DetectTags = "yes"
else ' if no category was selected
	var_piercing = ""
end if ' array for PIERCING TYPE selection


' Create array for THREADING selection
if request.querystring("threading") <> "" then

	threading_cleaned = request.querystring("threading")
	threading_cleaned = replace(threading_cleaned, "(", "")
	threading_cleaned = replace(threading_cleaned, ")", "")
	threading_cleaned = replace(threading_cleaned, "[", "")
	threading_cleaned = replace(threading_cleaned, "]", "")
	threading_cleaned = replace(threading_cleaned, """", "")

	'If more than one checkbox selected
	if Instr(threading_cleaned, ",") > 0 then
		
		threading_array = split(threading_cleaned, ",")
		var_build_threading = ""
		For i = 0 to Ubound(threading_array)
			if i = 0 then
				var_build_threading = var_build_threading +  replace(threading_array(i), " ", "+")
			else
				var_build_threading = var_build_threading + " or " + replace(threading_array(i), " ", "+")
			end if
		next
		var_threading = "(" + var_build_threading + ")"
	else ' if only one category selected
		var_threading = "(" + replace(threading_cleaned, " ", "+") + ")"
	end if
	
	if DetectTags = "yes" then
		var_threading = " and " + var_threading
	end if
	
	DetectTags = "yes"
else ' if no category was selected
	var_threading = ""
end if ' array for THREADING selection


' Create array for FLARE TYPE selection
if request.querystring("flare_type") <> "" then

	flare_type_cleaned = request.querystring("flare_type")
	flare_type_cleaned = replace(flare_type_cleaned, "(", "")
	flare_type_cleaned = replace(flare_type_cleaned, ")", "")
	flare_type_cleaned = replace(flare_type_cleaned, "[", "")
	flare_type_cleaned = replace(flare_type_cleaned, "]", "")
	flare_type_cleaned = replace(flare_type_cleaned, """", "")

	var_flares_fixed = Replace(flare_type_cleaned,"No", "non")
	'If more than one checkbox selected
	if Instr(flare_type_cleaned, ",") > 0 then
		
		flare_array = split(var_flares_fixed, ",")
		var_build_flare = ""
		For i = 0 to Ubound(flare_array)
			if i = 0 then
				var_build_flare = var_build_flare +  replace(flare_array(i), " ", "+")
			else
				var_build_flare = var_build_flare + " or " + replace(flare_array(i), " ", "+")
			end if
		next
		var_flare = "(" + var_build_flare + ")"
	else ' if only one category selected
		var_flare = "(" + replace(var_flares_fixed, " ", "+") + ")"
	end if
	
	if DetectTags = "yes" then
		var_flare = " and " + var_flare
	end if
	
	DetectTags = "yes"
else ' if no category was selected
	var_flare = ""
end if ' array for FLARE TYPE selection


' Create array for COLORS selection
if request.querystring("colors") <> "" then

	if request.querystring("color-filter") = "and" then
		var_color_filter = " and "
	else
		var_color_filter = " or "
	end if

	colors_cleaned = request.querystring("colors")
	colors_cleaned = replace(colors_cleaned, "(", "")
	colors_cleaned = replace(colors_cleaned, ")", "")
	colors_cleaned = replace(colors_cleaned, "[", "")
	colors_cleaned = replace(colors_cleaned, "]", "")
	colors_cleaned = replace(colors_cleaned, """", "")
	
	'If more than one checkbox selected
	if Instr(colors_cleaned, ",") > 0 then
		
		colors_array = split(colors_cleaned, ",")
		var_build_colors = ""
		For i = 0 to Ubound(colors_array)
			if i = 0 then
				var_build_colors = var_build_colors + replace(colors_array(i), " ", "+")
			else
				var_build_colors = var_build_colors + var_color_filter + replace(colors_array(i), " ", "+")
			end if
		next
		var_colors = "(" + var_build_colors + ")"
	else ' if only one category selected
		var_colors = "(" + replace(colors_cleaned, " ", "+") + ")"
	end if
	
	if detect_detail_tags = "yes" then
		var_colors = " and " + var_colors
	end if
	
	detect_detail_tags = "yes"
else ' if no category was selected
	var_colors = ""
end if ' array for COLORS selection



If request.querystring("discount") <> "" then

	discount_cleaned = request.querystring("discount")
	discount_cleaned = replace(discount_cleaned, "(", "")
	discount_cleaned = replace(discount_cleaned, ")", "")
	discount_cleaned = replace(discount_cleaned, "[", "")
	discount_cleaned = replace(discount_cleaned, "]", "")
	discount_cleaned = replace(discount_cleaned, """", "")

	if discount_cleaned = "5-20" then
		discount_build_string = "SaleDiscount:10 OR SaleDiscount:15 OR SaleDiscount:20"
	elseif discount_cleaned = "25-45" then
		discount_build_string = "SaleDiscount:25 OR SaleDiscount:30 OR SaleDiscount:35 OR SaleDiscount:40 OR SaleDiscount:45"
	elseif discount_cleaned = "50-70" then
		discount_build_string = "SaleDiscount:50 OR SaleDiscount:55 OR SaleDiscount:60 OR SaleDiscount:65 OR SaleDiscount:70"
	elseif discount_cleaned = "75-90" then
		discount_build_string = "SaleDiscount:75 OR SaleDiscount:80 OR SaleDiscount:85 OR SaleDiscount:90"
	elseif discount_cleaned = "all" then
		discount_build_string = "SaleDiscount:10 OR SaleDiscount:15 OR SaleDiscount:20 OR SaleDiscount:25 OR SaleDiscount:30 OR SaleDiscount:35 OR SaleDiscount:40 OR SaleDiscount:45 OR SaleDiscount:50 OR SaleDiscount:55 OR SaleDiscount:60 OR SaleDiscount:65 OR SaleDiscount:70 OR SaleDiscount:75 OR SaleDiscount:80 OR SaleDiscount:85 OR SaleDiscount:90"
	end if

	
	If DetectTags = "yes" then
		discount = " AND (" & discount_build_string & ")"
	else
	'	discount = "(SaleDiscount:" & discount_cleaned & ")"
		discount = "(" & discount_build_string & ")"
	end if
		DetectTags = "yes"
else
	discount = ""
end if

If request.querystring("limited") = "yes" then
	If DetectTags = "yes" then
		Limited = " and limited"
	else
		Limited = "limited"
	end if
		DetectTags = "yes"
else
	Limited = ""
end if


If request.querystring("onetime") = "yes" then
	If DetectTags = "yes" then
		OneTime = " and (onetime OR One+time+buy)"
	else
		OneTime = "(onetime OR One+time+buy)"
	end if
		DetectTags = "yes"
else
	OneTime = ""
end if

' Detect if pairs or singles
If request.querystring("pair") <> "" then
	if request.querystring("pair") = "pairs" then
		db_pair_sub = "pair"
	else
		db_pair_sub = "justone"
	end if
	
	If DetectTags = "yes" then
		var_pairs = " and (pair:" + db_pair_sub + ")"
	else
		var_pairs = "(pair:" + db_pair_sub + ")"
	end if
		DetectTags = "yes"
else
	var_pairs = ""
end if

If request.querystring("customorders") = "customorder-yes" or Instr(lcase(request.querystring("keywords")), "pre-order") or Instr(lcase(request.querystring("keywords")), "custom") or Instr(lcase(request.querystring("keywords")), "preorder") or Instr(lcase(request.querystring("keywords")), "pre order") or Instr(lcase(request.querystring("keywords")), "custom") or Instr(lcase(request.querystring("keywords")), "custom item") or Instr(lcase(request.querystring("keywords")), "custom order") then
	If DetectTags = "yes" then
		var_customorders = " and customorder-yes"
	else
		var_customorders = "customorder-yes"
	end if
		DetectTags = "yes"
elseif request.querystring("customorders") = "customorder-not" then
	If DetectTags = "yes" then
		var_customorders = " and customorder-not"
	else
		var_customorders = "customorder-not"
	end if
		DetectTags = "yes"		
else
	var_customorders = ""
end if
'response.write "PREORDERS: " & var_customorders


If request.querystring("filter-stock") = "" OR request.querystring("filter-stock") = NULL OR request.querystring("filter-stock") <> "all" then
	var_stock = NULL
Else
	var_stock = "all"
end if


If request.querystring("price") <> "select" AND request.querystring("price") <> "" AND request.querystring("price") <> "0;100" then
varprice = request.querystring("price")
else
varprice = NULL
end if

' NEW ------------------------
If request.querystring("new") = "Yes" then	
	SearchNew = "Yes"
else
	SearchNew = NULL
end if


If request.querystring("keywords") <> "" then

	Keywords = Sanitize(Request.querystring("keywords"))
	Keywords = Replace(LCase(Keywords),"  ", " ") ' replace accidentaly double space with single space
	Keywords = Replace(Keywords,"-", "")
	Keywords = Replace(Keywords,"]", "-")
	Keywords = Replace(Keywords,"[", "-")
	Keywords = Replace(Keywords,"(", "-")
	Keywords = Replace(Keywords,")", "-")


	
Function strClean (strtoclean)
Dim objRegExp, outputStr
Set objRegExp = New Regexp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "[(?*"",\\<>&#~%{}+_.@\!;]+"
outputStr = objRegExp.Replace(strtoclean, "-")

objRegExp.Pattern = "\-+"
outputStr = objRegExp.Replace(outputStr, "-")

objRegExp.Pattern = "\s{2,}" 'Removes duplicate spaces & condenses into one
outputStr = Trim(objRegExp.Replace(outputStr, " "))


strClean = outputStr
End Function



	cleaned_keywords = strClean(Keywords)
	'response.write cleaned_keywords

	'Change out words
	Keywords = Replace(cleaned_keywords,"belly ring", "navel") ' must come before single word
	Keywords = Replace(Keywords,"belly button", "navel")
	Keywords = Replace(Keywords,"belly", "navel")
	Keywords = Replace(Keywords,"eyebrow", "curved")
	Keywords = Replace(Keywords,"lip ring", "14g captive")
	Keywords = Replace(Keywords,"nipple ring", "captive")
	Keywords = Replace(Keywords,"tongue ring", "barbell")
	Keywords = Replace(Keywords,"tongue barbell", "barbell")
'	Keywords = Replace(Keywords,"nose ring", "captive 20g") ' modified below to search two categories
	Keywords = Replace(Keywords,"nostril nail", "nose hoop")
	Keywords = Replace(Keywords,"nose screw", "nosescrew")
	Keywords = Replace(Keywords,"nostril", "nosering")
	Keywords = Replace(Keywords,"captive ring", "captive")
	Keywords = Replace(Keywords,"cbr", "captive")
	Keywords = Replace(Keywords,"ball ends", "balls")
	Keywords = Replace(Keywords,"barbell ends", "balls")
	Keywords = Replace(Keywords,"hanger", "hanging")
	Keywords = Replace(Keywords,"ends", "balls")
	Keywords = Replace(Keywords," inch", "") ' added a space so it doesn't interfere with words that have inch in them (pincher)
	Keywords = Replace(Keywords,"jewelry", "-")
	Keywords = Replace(Keywords,"jewellry", "-")
	Keywords = Replace(Keywords,"piercing", "-")
	Keywords = Replace(Keywords,"tunnel", "eyelet")
	Keywords = Replace(Keywords,"tunnels", "eyelet")
	Keywords = Replace(Keywords,"dichro", "dichroic")
	Keywords = Replace(Keywords,"dichroicic", "dichroic")
	Keywords = Replace(Keywords,"little 7", "little7")
	Keywords = Replace(Keywords,"j curve", "j-curve")
	Keywords = Replace(Keywords,"oring", "orings")
	Keywords = Replace(Keywords,"oringss", "orings") ' remove duplicate
	Keywords = Replace(Keywords,"dring", "d-ring")
	Keywords = Replace(Keywords,"xring", "x-ring")
	Keywords = Replace(Keywords,"rose", "roses") ' Fixed error when inflectional searching of just wors "rose" gives results like "rise" to search
'	Keywords = Replace(Keywords,"per", "pincher") ' inch strips out inch in "pincher", have to search for per and then re-add the word pincher to work
	
	
	' THESARUS
	Keywords = Replace(Keywords,"7/16 gauge", "7/16 plug") 'for inch items
	Keywords = Replace(Keywords,"""", " plug") 'for inch items
	Keywords = Replace(Keywords,"g gauge", "g")
	Keywords = Replace(Keywords,"gauges", "plug")
	Keywords = Replace(Keywords," gauge", "g")
	Keywords = Replace(Keywords," ga", "g")
	Keywords = Replace(Keywords,"spacer", "plug")
	Keywords = Replace(Keywords,"spacers", "plug")
	
	'search for mis-pelled gauge
	Keywords = Replace(Keywords,"g guage", "g")
	Keywords = Replace(Keywords,"guages", "plug")
	Keywords = Replace(Keywords," guage", "g")
	
	' Keywords to overlook if typed in
	Keywords = Replace(Keywords," with ", " ")
	Keywords = Replace(Keywords,"custom", "")
	Keywords = Replace(Keywords,"pre-orders", "")
	Keywords = Replace(Keywords,"preorders", "")
	Keywords = Replace(Keywords,"pre orders", "")
	Keywords = Replace(Keywords,"preorder", "")
	Keywords = Replace(Keywords,"pre-order", "")
	Keywords = Replace(Keywords,"pre order", "")

	'Keywords = Replace(Replace(Trim(Request.querystring("keywords")), " AND ", " "), " ", " AND ")
	Keywords = Replace(Replace(Trim(Keywords), " AND ", " "), " ", " AND ")
	Keywords = Replace(Keywords,"""", "") 'remove double quotes
	Keywords = Replace(Keywords,"/", "slash") ' full text search doesn't like /, vw_product_search changes gauges with / to 1slash2 (1/2")
	
	Keywords = Replace(Keywords,"nose AND ring", "captive AND 20g OR nose")
	Keywords = Replace(Keywords,"dino", "dino OR dinosaur")
	Keywords = Replace(Keywords,"dinosaursaur", "dinosaur")
	Keywords = Replace(Keywords,"saddle", "saddle OR saddles")
	Keywords = Replace(Keywords,"saddles", "saddle OR saddles")
	
	' BANNER FIXES
	Keywords = Replace(Keywords,"junestone", "moonstones OR pearl")
	
	'Addition 20130730 to add FORMSOF(INFLECTIONAL,keyword) ajm
'	Keywords = "FORMSOF(INFLECTIONAL," + Keywords
'	Keywords = Keywords + ")"
'	Keywords = Replace(Keywords," AND ",") AND FORMSOF(INFLECTIONAL,")
'end 20130730 addition

'Addition 20130730 to add FORMSOF(INFLECTIONAL,keyword) ajm
if Keywords <> "" then
	Keywords = "FORMSOF(FREETEXT," + Keywords
	Keywords = Keywords + ")"
	Keywords = Replace(Keywords," AND ",") AND FORMSOF(FREETEXT,")
	Keywords = Replace(Keywords," OR ",") OR FORMSOF(FREETEXT,")
end if

	
	if InStr(Keywords, "http://") <> 0 then
	
		Response.End()
	End if
	'Keywords =  request.querystring("keywords")
else
	'Keywords = NULL 
end if

If var_category + var_brand + var_piercing + var_threading + var_flare + discount + Limited + OneTime + var_pairs + var_customorders <> "" then
		var_full_text_tags = var_category + var_brand + var_piercing + var_threading + var_flare + discount + Limited + OneTime + var_pairs + var_customorders
end if

If var_gauge + var_length + var_colors <> "" then
	var_full_text_detail_tags = var_gauge + var_length + var_colors
end if

'	SQL QUERY TO OUTPUT INFLECTIONAL FORMS OF A WORD
'	SELECT display_term, source_term, occurrence FROM sys.dm_fts_parser('FORMSOF(FREETEXT, "rose")', 1033, 0, 0)

'	response.write "Keywords: " + Keywords + "<br/>DB product build: " + var_full_text_tags + "<br/>DB details build: " + var_full_text_detail_tags + "<br/>Material build: " + var_material + "<br/>Querystring: " + Request.ServerVariables("QUERY_STRING") + "<br/>"

' If no querystring is found out, (user deletes all filters) then limit the results to the new items
If request.querystring() = "" then	
	SearchNew = "Yes"
	OrderBy = "new_page_date"
	DirectionOrder = "DESC"
end if

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandType = 4
objCmd.CommandText = "proc_search_flat_products"
objCmd.NamedParameters = True

If var_full_text_tags <> "" then
		objCmd.Parameters.Append(objCmd.CreateParameter("@tags",200,1,3000,var_full_text_tags))
end if

If var_full_text_detail_tags <> "" then
		objCmd.Parameters.Append(objCmd.CreateParameter("@detail_tags",200,1,3000,var_full_text_detail_tags))
end if

If(Instr(varprice, ";")>0) then
	arrPrice = split(request.querystring("price"), ";")
	If IsNumeric(arrPrice(0)) And IsNumeric(arrPrice(1)) Then
		objCmd.Parameters.Append(objCmd.CreateParameter("@price1",6,1,10,arrPrice(0))) 'Price range start
		objCmd.Parameters.Append(objCmd.CreateParameter("@price2",6,1,10,arrPrice(1))) 'Price range end
	End If
End If					



if Keywords <> "" then
	objCmd.Parameters.Append(objCmd.CreateParameter("@keywords",200,1,400,Keywords))
end if

If var_material <> "" then
		objCmd.Parameters.Append(objCmd.CreateParameter("@material",200,1,3000, var_material))
end if

If request.querystring("exclude-material") = "on" then
		objCmd.Parameters.Append(objCmd.CreateParameter("@materialexclude",200,1,5, "yes"))
else
	objCmd.Parameters.Append(objCmd.CreateParameter("@materialexclude",200,1,5, "no"))
end if

If request.querystring("feature") <> "" then
		objCmd.Parameters.Append(objCmd.CreateParameter("@feature",200,1,30, request.querystring("feature") ))
end if
		
		objCmd.Parameters.Append(objCmd.CreateParameter("@new",200,1,10,SearchNew))
		objCmd.Parameters.Append(objCmd.CreateParameter("@stock",200,1,10,var_stock))
		objCmd.Parameters.Append(objCmd.CreateParameter("@order",200,1,25,OrderBy))
		objCmd.Parameters.Append(objCmd.CreateParameter("@direction",200,1,10,DirectionOrder))
		
		If Session("ViewAll") = "yes" Then ' IF user wants all pages to be displayed as view all
			If Request.Querystring("RecordDisplay") <> "" Then ' If user wants to switch back to less items per page
				PageNumber = 1
				Session("ViewAll") = ""
			Else
				PageNumber = 1
			End If
		Else
			PageNumber = 1
		End if

		
		objCmd.Parameters.Append(objCmd.CreateParameter("@pagenumber",200,1,10,PageNumber))
		objCmd.Parameters.Append(objCmd.CreateParameter("@resultsperpage",200,1,10,50))
		
		If request.querystring("restock") = "restock" then
			objCmd.Parameters.Append(objCmd.CreateParameter("@restock",200,1,20,"yes"))	
		end if 
		' if admin user show all products active & inactive
		if request.cookies("adminuser") = "yes" and session("inactive") = "yes" then
			objCmd.Parameters.Append(objCmd.CreateParameter("@admin",200,1,20,"yes"))	
		end if
		

set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.Open objCmd

if NOT rsGetRecords.EOF then
	TotalRecords = rsGetRecords.RecordCount
end if

	if Request.form("resultsperpage") <> "" then
		Session("resultsperpage") = cint(Request.form("resultsperpage"))
	else
		if Session("resultsperpage") = "" then
			Session("resultsperpage") = 50
		end if
	end if


	' if view all is too high then reset back to 200 results per page
	if Session("resultsperpage") > 500 then
		Session("resultsperpage") = 200
	end if

if NOT rsGetRecords.EOF then
	if request.querystring("page") = "view-all" then ' for Google canonical link
		rsGetRecords.PageSize = 1000
	else ' default to regular session so a person doesn't come in from google and pull 1,000's of results on every page
		rsGetRecords.PageSize = Session("resultsperpage")
	end if	
	TotalPages = rsGetRecords.PageCount
end if

%>