<?xml version="1.0" encoding="ISO-8859-1"?>
<!--#include virtual="/Connections/sql_connection.asp" -->
<% 
Response.Buffer = true
Response.ContentType = "text/xml"

if request.querystring("q") = "" then
	sql = ""
elseif  request.querystring("q") = "glass-plugs" then
	sql = "material LIKE '%glass%' AND jewelry LIKE '%plugs%' AND j.type = 'None' AND  (brandname = 'Auxshine Jewelry' OR brandname = 'Very Clear Gems')"
elseif  request.querystring("q") = "new" then
	sql = "new_page_date >= GETDATE()-90"
end if

if sql <> "" then

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT p.ProductDetailID, j.ProductID, j.title, j.largepic, j.picture, j.picture_400, j.material, j.description, j.customorder, p.price, ShowTextLogo, brandname, SaleDiscount, flare_type, pair, gauge, length, img_full, ProductDetail1, qty, GaugeOrder, material FROM jewelry AS j INNER JOIN ProductDetails AS p ON j.ProductID = p.ProductID INNER JOIN dbo.TBL_Companies AS b ON j.brandname = b.name  LEFT OUTER JOIN tbl_images AS i ON p.img_id = i.img_id INNER JOIN TBL_GaugeOrder AS g ON p.Gauge = g.GaugeShow WHERE jewelry <> N'save' AND j.active = 1 AND p.active = 1 AND customorder <> 'yes' AND qty > 0 AND " + sql + " ORDER BY ProductID DESC, GaugeOrder ASC"

Set rsGetRecords = objCmd.Execute()

%>
<rss xmlns:g="http://base.google.com/ns/1.0" version="2.0">
  <channel>
    <title>Jewelry feed</title>
    <link>https://www.bodyartforms.com</link>
    <description>Bodyartforms jewelry feed</description>
<%

With rsGetRecords
Do While Not.Eof

sale_price = ""
pair = ""
var_brand = ""
detail_text = ""
image = ""
if rsGetRecords.Fields.Item("ShowTextLogo").Value <> "N" then
	var_brand = "<g:brand>" & rsGetRecords.Fields.Item("brandname").Value & "</g:brand>"
else
	var_brand = "<g:brand>Bodyartforms</g:brand>"
end if
If rsGetRecords.Fields.Item("SaleDiscount").Value > 0 then
	sale_price = "<g:sale_price>" & FormatNumber((rsGetRecords.Fields.Item("price").Value/100) * (100 - rsGetRecords.Fields.Item("SaleDiscount").Value), -1, -2, -2, -2) & "USD</g:sale_price>"
end if
If rsGetRecords.Fields.Item("pair").Value = "yes"  then
	pair = "Sold as a pair"
else
	pair = "Sold as a single"
end if
if rsGetRecords.Fields.Item("gauge").Value <> "" then
	detail_text = "Gauge: " & rsGetRecords.Fields.Item("gauge").Value
end if
if rsGetRecords.Fields.Item("length").Value <> "" then
	detail_text = detail_text & "Length: " & rsGetRecords.Fields.Item("length").Value
end if
if rsGetRecords("material") <> "" then
	detail_materials = detail_text & "  |  Materials: " & rsGetRecords("material")
end if
if rsGetRecords.Fields.Item("img_full").Value <> "" then
	image = "<g:image_link>https://bodyartforms-products.bodyartforms.com/" & rsGetRecords.Fields.Item("img_full").Value & "</g:image_link>"
else
	image = "<g:image_link>https://bafthumbs-400.bodyartforms.com/" & rsGetRecords.Fields.Item("picture_400").Value & "</g:image_link>"
end if

%>
    <item>
      <g:id><%= rsGetRecords.Fields.Item("ProductDetailID").Value %></g:id>
	  <g:title><%= left(rsGetRecords.Fields.Item("title").Value,100) %></g:title>
      <g:description><%= detail_text %>, <%= pair %>, <%= rsGetRecords.Fields.Item("flare_type").Value %><%= detail_materials %></g:description>
      <g:availability><% if rsGetRecords.Fields.Item("customorder").Value = "yes" then %>preorder<% else %>in stock<% end if %></g:availability>
	  <g:condition>new</g:condition>
	  <g:price><%= formatnumber(rsGetRecords.Fields.Item("price").Value,2) %> USD</g:price>
	  <%= sale_price %>
	  <g:link>http://www.bodyartforms.com/productdetails.asp?ProductID=<%= rsGetRecords.Fields.Item("ProductID").Value %></g:link>
	  <%= image %>
	  <%= var_brand %>
	  <g:size><%= detail_text %></g:size>
	  <g:color><%= rsGetRecords.Fields.Item("ProductDetail1").Value %></g:color>
	  <g:inventory><%= rsGetRecords.Fields.Item("qty").Value %></g:inventory>
	  <g:ordering_index><%= rsGetRecords.Fields.Item("GaugeOrder").Value %></g:ordering_index>
	  <g:age_group>adult</g:age_group>
	  <g:gender>unisex</g:gender>
	  <g:item_group_id><%= rsGetRecords.Fields.Item("ProductID").Value %></g:item_group_id>
	  <g:google_product_category>Apparel &amp; Accessories > Jewelry > Body Jewelry</g:google_product_category>
	  <g:product_type>Body Jewelry > Plugs</g:product_type>
    </item>
<%
.Movenext()
Loop
End With 


rsGetRecords.Close()
Set rsGetRecords = Nothing

end if ' if sql <> ""
%>
  </channel>
</rss>