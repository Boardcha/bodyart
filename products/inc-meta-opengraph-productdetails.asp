<% if not rsProduct.eof then %>

<!--Begin Facebook Open Graph protocol-->
<meta property="og:title" content="<%= meta_title %>"/>
<meta property="og:type" content="product"/>
<meta property="og:url" content="http://www.bodyartforms.com/productdetails.asp?ProductID=<%=(rsProduct.Fields.Item("ProductID").Value)%>"/>
<meta property="og:image" content="https://bodyartforms-products.bodyartforms.com/<%=(rsProduct.Fields.Item("picture").Value)%>"/>
<meta property="og:site_name" content="Bodyartforms"/>
<meta property="fb:app_id" content="180076978718781" />
<meta property="fb:page_id" content="149344708430326" />
<meta property="fb:admins" content="100000153477207" />
<meta property="og:description"
			content="<% If Not rsProduct.EOF Or Not rsProduct.BOF Then %><%=(rsProduct.Fields.Item("title").Value)%><% else %>Inactive product<% end if %>"/>
<!--End Facebook Open Graph protocol-->

<% end if ' not rsProduct.eof %>