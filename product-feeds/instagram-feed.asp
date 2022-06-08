<?xml version="1.0" encoding="ISO-8859-1"?>
<!--#include virtual="/Connections/sql_connection.asp" -->
<% 
' ======= EXTENSIVE FACEBOOK ARTICLE ON DEALING WITH INVENTORY PUSHES
' ======= https://developers.facebook.com/docs/marketplace/commerce-platform/inventory/
Response.Buffer = true
Response.ContentType = "text/xml"

if request.querystring("q") = "" then
	sql = ""
elseif request.querystring("q") = "brand" then
	sql = " AND b.searchable_brand_tags = '" & request.querystring("brand") & "'"
elseif request.querystring("q") = "new" then
	sql = " AND j.new_page_date >= GETDATE()-90"
elseif request.querystring("q") = "date_added" then
	sql = " AND YEAR(j.date_added) = YEAR('" & request.querystring("year") & "')"
	'sql = " AND j.new_page_date >= GETDATE()-90"
end if

if sql <> "" then

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT j.ProductID, ProductDetailID, j.title, ISNULL(d.Gauge,'') + ' ' + ISNULL(d.Length,'') + ' ' + ISNULL(d.ProductDetail1,'') AS 'variant_description', b.searchable_brand_tags, d.Gauge, j.picture, j.largepic, i.img_thumb, d.wearable_material, j.seo_meta_description, f.min_gauge, f.max_gauge, d.price, f.min_price, f.max_price, f.ShowTextLogo, ISNULL(flare_type,'') as flare_type, j.customorder, j.date_added, j.description, j.pair, REPLACE(REPLACE(TRIM(d.colors),'  ', ' '),' ' ,'/') as color_tags, j.jewelry, f.variants, d.qty AS 'inventory', CASE WHEN d.qty >=1 THEN 'in stock' ELSE 'out of stock' END AS 'availability', f.material, d.detail_materials FROM jewelry AS j INNER JOIN FlatProducts AS f ON j.ProductID = f.productid INNER JOIN ProductDetails as d ON f.productid = d.ProductID INNER JOIN dbo.TBL_Companies AS b ON j.brandname = b.name LEFT OUTER JOIN tbl_images AS i ON d.img_id = i.img_id WHERE j.jewelry <> N'save' AND j.active = 1 AND d.active = 1 AND j.customorder <> 'yes' AND wearable_material <> 'Acrylic' AND  wearable_material <> 'Bone' AND  wearable_material <> 'Horn' AND exclude_from_social_feeds = 0 AND j.title IS NOT NULL " + sql + " ORDER BY ProductID DESC"

Set rsGetRecords = objCmd.Execute()

%>
<rss xmlns:g="http://base.google.com/ns/1.0" version="2.0">
  <channel>
    <title>Jewelry feed</title>
    <link>https://www.bodyartforms.com</link>
    <description>Bodyartforms jewelry feed</description>
<%

While Not rsGetRecords.Eof

pair = ""
var_brand = ""
gauge_range = ""
price_range = ""

if rsGetRecords.Fields.Item("ShowTextLogo").Value <> "N" then
	var_brand = "<g:brand>" & rsGetRecords.Fields.Item("searchable_brand_tags").Value & "</g:brand>"
else
	var_brand = "<g:brand>Bodyartforms</g:brand>"
end if
If rsGetRecords.Fields.Item("pair").Value = "yes"  then
	pair_title = "PAIR "
	pair_description = " Sold as a pair "
else
	pair_title = "SINGLE "
	pair_description = " Sold as a single "
end if
if rsGetRecords.Fields.Item("min_gauge").Value = rsGetRecords.Fields.Item("max_gauge").Value then
	gauge_range = rsGetRecords.Fields.Item("min_gauge").Value
else
	gauge_range = rsGetRecords.Fields.Item("min_gauge").Value & " thru " & rsGetRecords.Fields.Item("max_gauge").Value
end if
if rsGetRecords.Fields.Item("min_price").Value = rsGetRecords.Fields.Item("max_price").Value then
	price_range = formatcurrency(rsGetRecords.Fields.Item("min_price").Value,2)
else
	price_range = formatcurrency(rsGetRecords.Fields.Item("min_price").Value,2) & " thru " & formatcurrency(rsGetRecords.Fields.Item("max_price").Value,2)
end if
if rsGetRecords("detail_materials") <> "" then
	detail_materials = "  |  Materials: " & Mid(replace(rsGetRecords("detail_materials"), "  ,", ","), 3) & "  |  "
end if
if rsGetRecords("wearable_material") <> "" then
	detail_wearable = "Wearable material: " & rsGetRecords("wearable_material") & "  |  "
end if

if rsGetRecords.Fields.Item("inventory").Value > 0 then
	availability = "in stock"
else
	availability = "out of stock"
end if

'===== google product category taxonomy =======
'===== 
if instr(rsGetRecords.Fields.Item("jewelry").Value, "necklace") > 0 then
	google_product_category = "196"
elseif  instr(rsGetRecords.Fields.Item("jewelry").Value, "earring") > 0 then
	google_product_category = "194"
elseif  instr(rsGetRecords.Fields.Item("jewelry").Value, "finger") > 0 then
	google_product_category = "200"
elseif  instr(rsGetRecords.Fields.Item("jewelry").Value, "bracelet") > 0 then
	google_product_category = "191"
elseif  instr(rsGetRecords.Fields.Item("jewelry").Value, "cleansers") > 0 then
	google_product_category = "2915"
else 
	google_product_category = "190"
end if
%>
	<% 
    '===== If it's a new product in the feed the show the id as the productid rather than the detail id 
        if rsGetRecords.Fields.Item("variants").Value <= 1 then
	%>
	<item>
		<g:id><%= rsGetRecords.Fields.Item("ProductDetailID").Value %></g:id>
		<g:title><%= pair_title & " " & rsGetRecords.Fields.Item("variant_description").Value & " #" & rsGetRecords.Fields.Item("ProductID").Value & " " & rsGetRecords.Fields.Item("flare_type").Value & " " & rsGetRecords.Fields.Item("title").Value %></g:title>
		<g:description><%= pair %><%= rsGetRecords.Fields.Item("variant_description").Value %><%= " " & rsGetRecords.Fields.Item("flare_type").Value & " " %><%= detail_materials %></g:description>
		<g:inventory><%= rsGetRecords.Fields.Item("inventory").Value %></g:inventory>
		<g:availability><%= availability %></g:availability>
		<g:condition>new</g:condition>
		<g:price><%= formatnumber(rsGetRecords.Fields.Item("min_price").Value,2) %> USD</g:price>
		<g:link>http://www.bodyartforms.com/productdetails.asp?ProductID=<%= rsGetRecords.Fields.Item("ProductID").Value %></g:link>
		<g:image_link>https://bodyartforms-products.bodyartforms.com/<%= rsGetRecords("largepic") %></g:image_link>
		<%= var_brand %>
		<g:size><%= rsGetRecords.Fields.Item("gauge").Value %></g:size>
		<g:age_group>adult</g:age_group>
		<g:gender>unisex</g:gender>
		<g:google_product_category><%= google_product_category %></g:google_product_category>
		<g:product_type><%= google_product_category %></g:product_type>
	</item>
    <% else ' ==== Item has variants, and IDs need to change %>
		<item>
            <g:id><%= rsGetRecords.Fields.Item("ProductDetailID").Value %></g:id>
			<g:item_group_id><%= rsGetRecords.Fields.Item("ProductID").Value %></g:item_group_id>
			<g:title><%= pair_title & " " & rsGetRecords.Fields.Item("variant_description").Value & " #" & rsGetRecords.Fields.Item("ProductID").Value & " " & rsGetRecords.Fields.Item("flare_type").Value & " " & rsGetRecords.Fields.Item("title").Value %></g:title>
			<g:description><%= rsGetRecords.Fields.Item("variant_description").Value %><%= pair_description %> <%= rsGetRecords.Fields.Item("flare_type").Value & " " %><%= detail_materials %><%= detail_wearable %></g:description>
			<g:inventory><%= rsGetRecords.Fields.Item("inventory").Value %></g:inventory>
			<g:availability><%= availability %></g:availability>
			<g:condition>new</g:condition>
			<g:price><%= formatnumber(rsGetRecords.Fields.Item("price").Value,2) %> USD</g:price>
			<g:link>http://www.bodyartforms.com/productdetails.asp?ProductID=<%= rsGetRecords.Fields.Item("ProductID").Value %></g:link>
			<g:image_link>https://bodyartforms-products.bodyartforms.com/<%= rsGetRecords("largepic") %></g:image_link>
			<%= var_brand %>
			<g:size><%= rsGetRecords.Fields.Item("gauge").Value %></g:size>
			<g:color><%= rsGetRecords.Fields.Item("color_tags").Value %></g:color>
			<g:age_group>adult</g:age_group>
			<g:gender>unisex</g:gender>
			<g:google_product_category><%= google_product_category %></g:google_product_category>
			<g:product_type><%= google_product_category %></g:product_type>
		</item>
    <% end if '==== variants = 1 %>   
<%

rsGetRecords.Movenext()
wend


rsGetRecords.Close()
Set rsGetRecords = Nothing

end if ' if sql <> ""
%>
  </channel>
</rss>