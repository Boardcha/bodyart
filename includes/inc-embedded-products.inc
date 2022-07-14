<div class="mt-2 full-width-block">
	<div class="baf-carousel embedded-products">
				<% 

				SqlString = SqlString 
				Set rsGetEmbedded = DataConn.Execute(SqlString)


				i = 1
					While NOT rsGetEmbedded.EOF 

					' set variables for pricing
					if rsGetEmbedded.Fields.Item("min_sale_price").Value <> "" then
					min_price = FormatNumber(rsGetEmbedded.Fields.Item("min_sale_price").Value,2)
					else
					min_price = ""
					end if
					if rsGetEmbedded.Fields.Item("max_sale_price").Value <> "" then
					max_price = FormatNumber(rsGetEmbedded.Fields.Item("max_sale_price").Value,2)
					else
					max_price = ""
					end if

					DisplayPrice = ""
					if rsGetEmbedded.Fields.Item("SaleDiscount").Value > 0 then	
					DisplayPrice = DisplayPrice & rsGetEmbedded.Fields.Item("SaleDiscount").Value & "% OFF "
					end if


					if rsGetEmbedded.Fields.Item("min_sale_price").Value <> "" then
					DisplayPrice = DisplayPrice & "$" & min_price & " "
					end if

					if rsGetEmbedded.Fields.Item("min_sale_price").Value <> rsGetEmbedded.Fields.Item("max_sale_price").Value then
					DisplayPrice = DisplayPrice & " - $" & max_price
					end if
					%>
						<a class="slide text-dark homepage-graphic" href="/productdetails.asp?productid=<%= rsGetEmbedded.Fields.Item("ProductID").Value %>" id="new-<%= replace(lcase(rsGetEmbedded.Fields.Item("title").Value)," ", "-") %>-<%=(rsGetEmbedded.Fields.Item("ProductID").Value)%>">
							<% if i < 9 then %>
							<img class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetEmbedded.Fields.Item("picture").Value %>" alt="<%=(rsGetEmbedded.Fields.Item("title").Value)%>" title="<%=(rsGetEmbedded.Fields.Item("title").Value)%>">
							<% else %><!-- lazy load in images beyond 8-->
							<img class="img-fluid lazyload" src="/images/image-placeholder.png" data-src="https://bafthumbs-400.bodyartforms.com/<%= rsGetEmbedded.Fields.Item("picture").Value %>" alt="<%=(rsGetEmbedded.Fields.Item("title").Value)%>" title="<%=(rsGetEmbedded.Fields.Item("title").Value)%>">
							<% end if 
							i = i +1 %>
						<div class="w-100 text-center px-1">
								<div class="small">
										<%= rsGetEmbedded.Fields.Item("title").Value %>
								</div>
								<div class="small font-weight-bold  d-block">
										<%= DisplayPrice %>
								</div>
								<div class="small font-weight-bold  d-block">
										<%= rsGetEmbedded.Fields.Item("min_gauge").Value %>
										<% if rsGetEmbedded.Fields.Item("min_gauge").Value <> rsGetEmbedded.Fields.Item("max_gauge").Value then %> 
										- <%= rsGetEmbedded.Fields.Item("max_gauge").Value %>
										<% end if %>
								</div>
							
							</div> 
					</a>
					<% 
rsGetEmbedded.MoveNext()
Wend
%>
				</div>
</div><!-- full-width-block -->