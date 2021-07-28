<%@LANGUAGE="VBSCRIPT" %>
<% response.Buffer=false %> 
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT * FROM sent_items  WHERE ship_code = 'paid' AND (Review_OrderError <> 1 OR  Review_OrderError IS NULL) AND shipped = 'Pending shipment' AND shipping_type LIKE '%DHL%' ORDER BY PackagedBy, CASE WHEN shipping_type LIKE '%office%' THEN 1 WHEN autoclave = 1 THEN 2 WHEN shipping_type LIKE '%express%' THEN 4 WHEN shipping_type LIKE '%ups%' THEN 5 WHEN shipping_type LIKE '%USPS Priority mail%' THEN 6 WHEN (sent_items.shipping_type = 'USPS First Class Mail') THEN 7 WHEN (sent_items.shipping_type LIKE '%global basic%') THEN 8 WHEN shipping_type LIKE '%max%' THEN 9 WHEN   (sent_items.shipping_type = 'DHL Basic mail') THEN 10 WHEN  (sent_items.shipping_type = 'DHL GlobalMail Packet Priority') THEN 11 WHEN  (sent_items.shipping_type = 'DHL GlobalMail Parcel Priority') THEN 12 ELSE 20 END ASC, ID ASC"
Set rsGetInvoice = objCmd.Execute()
%>
<html class="print-invoice">
<head>
<title>INVOICE</title>
<link rel="stylesheet" href="/CSS/4001dd09.css?v=102016">
<link href="/CSS/print-friendly.css" rel="stylesheet" type="text/css">
<link href='https://fonts.googleapis.com/css?family=Economica' rel='stylesheet'>
</head>
<body>

<%

  Function CleanUp (input)
      Dim objRegExp, outputStr
      Set objRegExp = New Regexp

      objRegExp.IgnoreCase = True
      objRegExp.Global = True
      objRegExp.Pattern = "((?![a-zA-Z]).)+"
      outputStr = objRegExp.Replace(input, "-")

      objRegExp.Pattern = "\-+"
      outputStr = objRegExp.Replace(outputStr, " ")

      CleanUp = outputStr
    End Function

    

page_num = 1
While NOT rsGetInvoice.EOF


if shipping_type <> rsGetInvoice.Fields.Item("shipping_type").Value and rsGetInvoice.Fields.Item("autoclave").Value <> 1 then
%>
<div style="text-align: center;font-size:2em;font-weight:bold; padding-top: 12em;">
<%= CleanUp(rsGetInvoice.Fields.Item("shipping_type").Value)  %>
</div>
<div class="page-break"></div>
<%
end if

shipping_type = rsGetInvoice.Fields.Item("shipping_type").Value
%>
<div class="print-wrapper">
<table class="header">
  <tr>
    <td><img src="/images/bodyartforms-solid-text.png" class="logo-text"><br>
      1966 S. Austin Ave.<br>
        Georgetown, TX  78626<br>
        service@bodyartforms.com<br/>
		(877) 223-5005</td>
    <td>
	<div class="barcode">
	<img src="barcode.asp?code=<%=(rsGetInvoice.Fields.Item("ID").Value)%>&height=30&width=1&mode=code39&text=0"><br><span style="font-size:18px;font-weight:bold"><%= page_num %> - <%=(rsGetInvoice.Fields.Item("ID").Value)%></span>
	<div>
		<% if instr(shipping_type, "DHL") > 0 then %>
		<img src="../images/dhl-svg.svg">
		<% end if %>
		<% if instr(shipping_type, "USPS") > 0 then %>
		<img src="../images/usps-svg.svg">
		<% end if %>
		<% if instr(shipping_type, "UPS") > 0 then %>
		<img src="../images/ups-svg.svg">
		<% end if %>
	</div>
        <% if (rsGetInvoice.Fields.Item("PackagedBy").Value) <> "" then %>
        Packaged by: <%=(rsGetInvoice.Fields.Item("PackagedBy").Value)%>
        <% end if %>
    </div>
	<div class="customer-address">
        <span style="font-size:18px;font-weight:bold"><%= rsGetInvoice.Fields.Item("customer_first").Value %>&nbsp;<%=(rsGetInvoice.Fields.Item("customer_last").Value)%></span><% if (rsGetInvoice.Fields.Item("company").Value) <> "" then %><br>
<%=(rsGetInvoice.Fields.Item("company").Value)%><% end if %><br>
        <%=(rsGetInvoice.Fields.Item("address").Value)%> <br>
        <% if (rsGetInvoice.Fields.Item("address2").Value) <> "" then %>
        <%=(rsGetInvoice.Fields.Item("address2").Value)%> <br>
        <% end if %>
        <%=(rsGetInvoice.Fields.Item("city").Value)%>, <%=(rsGetInvoice.Fields.Item("state").Value)%><%=(rsGetInvoice.Fields.Item("province").Value)%>&nbsp;&nbsp; <%=(rsGetInvoice.Fields.Item("zip").Value)%><br>
    <%=(rsGetInvoice.Fields.Item("country").Value)%>
		</div>
	</td>
  </tr>
</table>

  <%

Set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT title FROM dbo.QRY_OrderDetails WHERE ID = ? ORDER BY ID_Description, location, ProductDetailID"
objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,10,rsGetInvoice.Fields.Item("ID").Value))
Set rsGetOrderDetails2 = objCmd.Execute()

While Not rsGetOrderDetails2.Eof

if InStr( 1, (rsGetOrderDetails2.Fields.Item("title").Value), "RETURN MAILER", vbTextCompare) then
	ReturnMailer = "yes"
end if

rsGetOrderDetails2.Movenext()
Wend

%>
<% if rsGetInvoice.Fields.Item("pay_method").Value = "Etsy" then %>
<center><h1>ETSY</h1></center>
<% end if %>
 
<table>
<td>
<table class="items">
<% if rsGetInvoice.Fields.Item("item_description").Value <> "" then %>
	<td class="public-notes" colspan="8">
		<%=rsGetInvoice.Fields.Item("item_description").Value %>
	</td>
<% end if

Set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT OrderDetailID, TBL_OrderSummary.qty, title, ProductDetail1, PreOrder_Desc, ProductDetails.price, notes, TBL_OrderSummary.item_price, DetailCode, ProductDetailID, ID_Description, ID_SortOrder, ID_Number, type, autoclavable, BinNumber_Detail, Gauge, Length, location FROM TBL_OrderSummary INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID INNER JOIN sent_items ON TBL_OrderSummary.InvoiceID = sent_items.ID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE ID = ? ORDER BY ID_SortOrder ASC, BinNumber_Detail ASC, ProductDetailID ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,10,rsGetInvoice.Fields.Item("ID").Value))
Set rsGetOrderDetails = objCmd.Execute()

LineItem = 0
SumLineItem = 0

ItemsTotal = 0
%>
  <% While Not rsGetOrderDetails.Eof %>
  <tr>
    <td class="line-qty">
		<%=(rsGetOrderDetails.Fields.Item("qty").Value)%>
	</td>
    <td>
		<% if rsGetOrderDetails.Fields.Item("notes").Value <> "" then %>
			<strong><%= rsGetOrderDetails.Fields.Item("notes").Value %> &nbsp;</strong>
		<% end if %>
				<% if rsGetOrderDetails.Fields.Item("autoclavable").Value = 1 and rsGetInvoice.Fields.Item("autoclave").Value = 1 then %>
			<strong>
		<% end if %>
		<%=(rsGetOrderDetails.Fields.Item("title").Value)%>
				<% if rsGetOrderDetails.Fields.Item("autoclavable").Value = 1 then %>
			</strong>
		<% end if %>
		&nbsp;&nbsp;<%=(rsGetOrderDetails.Fields.Item("ProductDetail1").Value)%>&nbsp;<%=(rsGetOrderDetails.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetOrderDetails.Fields.Item("Length").Value)%>
        <% if InStr( 1, (rsGetOrderDetails.Fields.Item("title").Value), "PRE-ORDER", vbTextCompare) then %>
        <br>
        <%= (rsGetOrderDetails.Fields.Item("PreOrder_Desc").Value) %>
        <% end if %>
		
		<span class="location">
		  <%= rsGetOrderDetails.Fields.Item("ID_Description").Value %>
		  <% if rsGetOrderDetails.Fields.Item("BinNumber_Detail").Value <> 0 then %>
		  LIM BIN <%= rsGetOrderDetails.Fields.Item("BinNumber_Detail").Value %>&nbsp;&nbsp;
		  <% end if %>
	&nbsp;      <%=(rsGetOrderDetails.Fields.Item("location").Value)%>
		</span>
		
    </td>
	<% if InStr(rsGetInvoice.Fields.Item("item_description").Value, "GIFT") < 1 then %>
    <td class="line-cost">
        <%= FormatCurrency((rsGetOrderDetails.Fields.Item("item_price").Value), -1, -2, -0, -2) %>
    </td>
    <td class="line-total">
        <%= FormatCurrency((rsGetOrderDetails.Fields.Item("item_price").Value)*(rsGetOrderDetails.Fields.Item("qty").Value), -1, -2, -0, -2) %>
    </td>
	<% end if %>
  </tr>
  <%
	LineItem = rsGetOrderDetails.Fields.Item("item_price").Value * rsGetOrderDetails.Fields.Item("qty").Value
	
	SumLineItem = SumLineItem + LineItem
	rsGetOrderDetails.Movenext()
	InvoiceTotal = SumLineItem + (rsGetInvoice.Fields.Item("shipping_rate").Value) - (rsGetInvoice.Fields.Item("coupon_amt").Value)
Wend

Set rsGetOrderDetails = Nothing
%>
</table>
</td>
</table>

<table class="table-wrapper">
<td class="thank-you-area">
	<div class="thank-you">
		THANK YOU for ordering with us!
	</div>
	<div class="our-promise">
		We stand behind our products and service...and we want you to be 100% happy with your purchase. If you aren't, we're here to help make it right. 
	</div>
	<div class="thank-you">
		Text BODYART to 22828 to get notified about sales via our newsletter
	</div>
</td>
<td class="wrapper-totals">
	<table class="totals">
	<% if InStr(rsGetInvoice.Fields.Item("item_description").Value, "GIFT") < 1 then %>
	  <tr>
		<td class="line-cost">Subtotal</td>
		<td class="line-total">
			<%= FormatCurrency(SumLineItem, -1, -2, -0, -2)%>
		</td>
	  </tr>
	 <%

	' Array for invoice totals
	ReDim arrTotals(2,5) 

	'arrTotals(col,row)
	arrTotals(0,0) = "10% preferred discount" 
	arrTotals(1,0) = "total_preferred_discount" 
	total_preferred_discount = rsGetInvoice.Fields.Item("total_preferred_discount").Value
	arrTotals(2,0) = "&#8722;"
	arrTotals(0,1) = "Coupon discount" 
	arrTotals(1,1) = "total_coupon_discount" 
	total_coupon_discount = rsGetInvoice.Fields.Item("total_coupon_discount").Value
	arrTotals(2,1) = "&#8722;" 
	if rsGetInvoice.Fields.Item("country").Value = "GB" OR rsGetInvoice.Fields.Item("country").Value = "Great Britain" OR rsGetInvoice.Fields.Item("country").Value = "Great Britain and Northern Ireland" OR rsGetInvoice.Fields.Item("country").Value = "United Kingdom" then
		arrTotals(0,2) = "VAT" 
	else
		arrTotals(0,2) = "Tax"
	end if
	arrTotals(1,2) = "total_sales_tax" 
	total_sales_tax = rsGetInvoice.Fields.Item("total_sales_tax").Value
	arrTotals(2,2) = "&nbsp;&nbsp;"
	arrTotals(0,3) = "Gift certificate" 
	arrTotals(1,3) = "total_gift_cert"
	total_gift_cert = rsGetInvoice.Fields.Item("total_gift_cert").Value 
	arrTotals(2,3) = "&#8722;"
	arrTotals(0,4) = "Free gift (USE NOW) credits" 
	arrTotals(1,4) = "total_free_credits" 
	total_free_credits = rsGetInvoice.Fields.Item("total_free_credits").Value
	arrTotals(2,4) = "&#8722;"
	arrTotals(0,5) = "Store account credit" 
	arrTotals(1,5) = "total_store_credit"
	total_store_credit = rsGetInvoice.Fields.Item("total_store_credit").Value
	arrTotals(2,5) = "&#8722;"


	For i = 0 to UBound(arrTotals, 2) 

		if rsGetInvoice.Fields.Item(arrTotals(1,i)).Value <> 0 then
	%>
	  <tr>
		<td class="line-cost"><%= arrTotals(0,i) %></td>
		<td class="line-total"><%= arrTotals(2,i) %>
		<%= FormatCurrency(rsGetInvoice.Fields.Item(arrTotals(1,i)).Value, -1, -2, -0, -2) %>
		</td>
	  </tr>
	<% 
		end if ' if i > 2 or values not 0
	next ' loop through totals array


	InvoiceTotal = (SumLineItem - total_preferred_discount - total_coupon_discount - total_free_credits + rsGetInvoice.Fields.Item("shipping_rate").Value + total_sales_tax - total_store_credit - total_gift_cert)
	%>
	<% end if ' if gift order don't show pricing BUT DO show shipping method below 
	%>
	  <tr>
		<td class="line-cost"><%= (rsGetInvoice.Fields.Item("shipping_type").Value) %></td>
		<td class="line-total"><%= FormatCurrency((rsGetInvoice.Fields.Item("shipping_rate").Value), -1, -2, -0, -2)%></td>
	  </tr>
	  <% if InStr(rsGetInvoice.Fields.Item("item_description").Value, "GIFT") < 1 then %>
	  <tr>
		<td class="line-cost bold">
		  TOTAL
		</td>
		<td class="line-total bold"> <% if InvoiceTotal < 0 then %>0<% else %><%= FormatCurrency(InvoiceTotal, -1, -2, -0, -2) %><% end if %>
		   
		</td>
	  </tr>
	  <%
	  end if ' if gift order do not display totals 
	  %>
	</td>
</table>
</table>
<% 
if ReturnMailer <> "yes" then %>

<div class="policies">
	<strong>Something is wrong with my order. What do I do?</strong>
	<br/>
	First off, we're very sorry that something was wrong on your order! And we'll work as quickly as possible to fix it. Any item that has arrived damaged or incorrect is eligible for a replacement at no additional shipping cost, even if it's not in its sealed bag.
	<br/>
	<br/><strong>Returns</strong>
	<br/>
	We want you to be 100% satisfied with your order. We take returns up to 30 days after you receive your order.
	<br/>
	
		<ul>
			<li>
			For safety, ALL body jewelry &amp; earrings come in a sealed baggie. As long as it is still in the untampered sealed baggie we can take it back. Please be sure to measure and inspect all jewelry before breaking the seal.
			</li>
			<li>Shipping costs for returns are the responsibility of the customer unless the item(s) arrived damaged or incorrect. </li>
			<li>
			Necklaces, finger rings, bracelets, & clothing can be return for ANY reason, even out of the sealed baggies, as long as the item is still in its original condition. 
			</li>
		</ul>
<strong>For any and all concerns, returns, questions, or comments, you can contact us at service@bodyartforms.com</strong>		
</div>

<% else 'DOES have a return mailer 
%>
<div class="policies">
	<strong>What do I do with my incorrect items?</strong>
	<br/>
	<ul>
		<li>
			If you DID receive a return mailer,  please send the defective/wrong product(s) back to us in the provided mailer (postage provided).
		</li>
		<li>
			If you did NOT receive a return envelope with this shipment then don't worry about sending your items back.
		</li>
		</ul>
</div>
  <% end if ' DOES have a return mailer 
%>

</div>
<div class="page-break"></div>
 <% 
 ReturnMailer = "" ' this MUST be here otherwise an order will have a return mailer and each order after that will say the return mailer message.

page_num = page_num + 1
  rsGetInvoice.MoveNext()
Wend
%>
</body>
</html>
<%
rsGetInvoice.Close()
%>
