        <%
                if var_meta_productdetails = "yes" then 
                set objCmd = Server.CreateObject("ADODB.command")
                objCmd.ActiveConnection = DataConn
                objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetailID, ProductID, detail_code, price, qty, ISNULL(Gauge, '') + ' ' + ISNULL(Length, '') + ' ' + ISNULL(ProductDetail1, '') AS offer_name FROM ProductDetails WHERE  (ProductID = ?) ORDER BY item_order, price"
                objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,ProductID))
                
                
                set rs_getOffers = Server.CreateObject("ADODB.Recordset")
                rs_getOffers.CursorLocation = 3 'adUseClient
                rs_getOffers.Open objCmd
                total_item_offers = rs_getOffers.RecordCount
        
                if not rsProduct.eof then
                %>
        <link rel="canonical" href="https://bodyartforms.com/productdetails.asp?productid=<%= rsProduct.Fields.Item("ProductID").Value %>" />
        <!--#include virtual="/products/inc-meta-opengraph-productdetails.asp" -->
        <% end if %>
        <% If not rsProduct.eof then

        Public Function ToIsoDate(datetime)
            ToIsoDate = Year(datetime) & "-" & Month(datetime) & "-" & Day(datetime)
        End Function  
        %>
        <script type="application/ld+json">
                        {
                          "@context": "http://schema.org/",
                          "@type": "Product",
                          "name": "<%= replace(meta_title, """", " inch") %>",
                          "image": "https://bodyartforms-products.bodyartforms.com/<%=(rsProduct.Fields.Item("largepic").Value)%>",
                          "description": "<%= replace(meta_description, """", " inch") %>",
                          "sku": "<%= rsProduct.Fields.Item("ProductID").Value %>",
                          "brand": {
                                "@type": "Thing",
                                "name": "<%= rsProduct.Fields.Item("brandname").Value %>"
                          },
                        <% if var_show_star_ratings = "yes" then %> 
                          "aggregateRating": {
                                "@type": "AggregateRating",
                                "ratingValue": "<%= avg_rating %>",
                                "reviewCount": "<%= total_ratings %>"
                          },
                        <% if NOT rsJsonReviews.eof then %>
                        "review": [
                        <% i = 1 
                        While NOT rsJsonReviews.EOF 
                        %>
                        {
                          "@type": "Review",
                          "reviewRating": {
                            "@type": "Rating",
                            "ratingValue": "<%= rsJsonReviews.Fields.Item("review_rating").Value %>"
                          },
                          "author": {
                            "@type": "Person",
                            "name": "<%= rsJsonReviews.Fields.Item("name").Value %>"
                          },
                          "reviewBody": "<%= replace(rsJsonReviews.Fields.Item("review").Value, """", "") %>"
                        }<% if i <> total_json_reviews then %>,<%
                      end if ' only show comma if it's not the last record
                        i = i + 1
                        rsJsonReviews.movenext()
                        wend %>],<% end if   'NOT rsJsonReviews.eof 
                        end if '==== var_show_star_ratings = yes
                        %>
                          "offers": [
                                  <% 
                                  i = 1
                                  while not rs_getOffers.eof 

                                    if rs_getOffers.Fields.Item("qty").Value > 0 then
                                      offer_item_availability = "InStock"
                                    end if
                                    if rsProduct.Fields.Item("type").Value = "One time buy" OR rsProduct.Fields.Item("type").Value = "limited" then
                                       offer_item_availability = "LimitedAvailability"
                                    end if
                                    if rsProduct.Fields.Item("type").Value = "Discontinued" then
                                       offer_item_availability = "Discontinued"
                                    end if
                                    if rs_getOffers.Fields.Item("qty").Value <= 0 then
                                      offer_item_availability = "OutOfStock"
                                    end if
                                    if rsProduct.Fields.Item("customorder").Value = "yes" then
                                    offer_item_availability = "PreOrder"
                                    end if
                                  %>
                                  
                                  {
                                        "@type": "Offer",
                                        "availability": "http://schema.org/<%= offer_item_availability %>",
                                        "price": "<%= FormatNumber(rs_getOffers.Fields.Item("price").Value,2) %>",
                                        "priceValidUntil": "<%= ToIsoDate(now() + 365) %>",
                                        "priceCurrency": "USD",
                                        <% if rs_getOffers.Fields.Item("detail_code").Value <> "" then %>
                                          "mpn": "<%= rs_getOffers.Fields.Item("detail_code").Value %>",
                                        <% end if %>
                                        "itemCondition": "http://schema.org/NewCondition",
                                        "url": "http://www.bodyartforms.com/productdetails.asp?ProductID=<%= rsProduct.Fields.Item("ProductID").Value %>",
                                           "itemOffered" : 
                                                {
                                                        "@type" : "Thing",
                                                        "description" : "<%= replace(rs_getOffers.Fields.Item("offer_name").Value, """", " inch") %>",
                                                        "name" : "<%= replace(rs_getOffers.Fields.Item("offer_name").Value, """", " inch") & " " %><%= replace(meta_title, """", " inch") %>"
                                                }
                                  }<% if i <> total_item_offers then %>,
                                  <%
                                end if ' only show comma if it's not the last record
                                  i = i + 1
                                  rs_getOffers.movenext()
                                  wend %>
                                ]
                        }
                </script>
        <%  end if ' not rsProduct.eof %>
        <% end if %>
        <% if var_meta_products_aggregate = "yes" then 
        if Request.ServerVariables("QUERY_STRING") <> "" then
        if NOT rsSiteMap.eof then
        if rsSiteMap.Fields.Item("canonical_url").Value <> "" then 
        %>
                <link rel="canonical" href="https://bodyartforms.com/products.asp?<%= rsSiteMap.Fields.Item("canonical_url").Value %>" />
        <% end if
        end if 
        end if %>
                <% If CurrentPage > 1 then %>
                <link rel="prev" href="https://bodyartforms.com/products.asp?<%= Replace(var_qs_url, "&pagenumber=" & CurrentPage, "") %>&pagenumber=<%= CurrentPage - 1 %>">
                <%
                End if

                If Cint(CurrentPage) < Cint(TotalPages) then
                %>
                <link rel="next" href="https://bodyartforms.com/products.asp?<%= Replace(var_qs_url, "&pagenumber=" & CurrentPage, "") %>&pagenumber=<%= CurrentPage + 1 %>">
        <% end if %>
        <meta name="revisit-after" content="15 days" />
        <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
        <% end if %>
        <% if var_extra_head_inc = "homepage" then %>
        <script type="application/ld+json">
                        {
                          "@context": "http://schema.org",
                          "@type": "JewelryStore",
                          "@id": "https://www.bodyartforms.com",
                          "name": "Bodyartforms",
                          "logo": "http://www.bodyartforms.com/images/baf-logo-head-text.png",
                          "image": "http://www.bodyartforms.com/images/baf-logo-head-text.png",
                          "priceRange": "$$",
                          "potentialAction": {
                            "@type": "SearchAction",
                            "target": "https://www.bodyartforms.com/products.asp?keywords={search_term_string}",
                            "query-input": "required name=search_term_string"
                          },
                          "telephone": "+18772235005",
                          "url": "http://www.bodyartforms.com",
                            "contactPoint": [{
                                "@type": "ContactPoint",
                                "telephone": "+1-877-223-5005",
                                "contactType": "Customer service"
                          }],
                            "address": {
                                "@type": "PostalAddress",
                                "streetAddress": "1966 S Austin Ave",
                                "addressLocality": "Georgetown",
                                "addressRegion": "TX",
                                "postalCode": "78626",
                                "addressCountry": "US"
                          },
                          "geo": {
                                "@type": "GeoCoordinates",
                                "latitude": 30.626,
                                "longitude": -97.679081
                          },
                          "openingHoursSpecification": [
                          {
                                "@type": "OpeningHoursSpecification",
                                "dayOfWeek": [
                                  "Monday",
                                  "Tuesday",
                                  "Wednesday",
                                  "Thursday",
                                  "Friday"
                                ],
                                "opens": "09:00",
                                "closes": "17:00"
                          }
                        ],
                          "sameAs": [
                                "https://www.facebook.com/pages/Bodyartforms/149344708430326",
                                "http://instagram.com/bodyartforms",
                                "https://www.pinterest.com/bodyartforms/",
                                "https://plus.google.com/+bodyartforms/posts"
                          ]
                        }
                </script>
        <% end if %>