		<% '==== only show main first product when i=0, otherwise show variant id's
		if var_temp_productid <> rsGetRecords.Fields.Item("ProductID").Value then
			i = 0 
		end if 
        if i = 0 then
		  
		
		'<item>
			'<g:id><%= rsGetRecords.Fields.Item("ProductID").Value %></g:id>
			'<g:item_group_id><%= rsGetRecords.Fields.Item("ProductID").Value %></g:item_group_id>
			'<g:title><%= rsGetRecords.Fields.Item("title").Value & " #" & rsGetRecords.Fields.Item("ProductID").Value %></g:title>
			'<g:description><%= price_range %>, <%= gauge_range %> <%= pair_description %> <%= rsGetRecords.Fields.Item("flare_type").Value %><%= rsGetRecords.Fields.Item("seo_meta_description").Value %></g:description>
			'<g:availability><%= rsGetRecords.Fields.Item("availability").Value %></g:availability>
			'<g:condition>new</g:condition>
			'<g:price><%= formatnumber(rsGetRecords.Fields.Item("price").Value,2) %> USD</g:price>
			'<g:link>http://www.bodyartforms.com/productdetails.asp?ProductID=<%= rsGetRecords.Fields.Item("ProductID").Value %></g:link>
			'<g:image_link>https://bafthumbs-400.bodyartforms.com/<%=  rsGetRecords.Fields.Item("picture").Value %></g:image_link>
			'<%= var_brand %>
			'<g:size><%= gauge_range %></g:size>
			'<g:color><%= rsGetRecords.Fields.Item("color_tags").Value %></g:color>
			'<g:age_group>adult</g:age_group>
			'<g:gender>unisex</g:gender>
			'<g:google_product_category><%= google_product_category %></g:google_product_category>
			'<g:product_type><%= google_product_category %></g:product_type>
		'</item>

		 i = i + 1
		end if %>