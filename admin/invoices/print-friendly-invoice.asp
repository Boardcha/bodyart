<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsGetInvoice__MMColParam
if (Request.form("invoice_num") <> "") then 
rsGetInvoice__MMColParam = Request.form("invoice_num") 
else
rsGetInvoice__MMColParam = Request.QueryString("ID")
end if
%>
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT *  FROM sent_items  WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("invoice",3,1,10,rsGetInvoice__MMColParam))
Set rsGetInvoice = objCmd.Execute()

' ---- detect gift order
if InStr(rsGetInvoice.Fields.Item("item_description").Value, "GIFT") >= 1 then
	var_gift_order = "yes"
end if
%>



<html class="print-invoice">
<head>
<title>INVOICE</title>
<link rel="stylesheet" href="/CSS/4001dd09.css?v=102016">
<link href="/CSS/print-friendly.css" rel="stylesheet" type="text/css">
<link href='https://fonts.googleapis.com/css?family=Economica' rel='stylesheet'>
</head>
<body>
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
	<img src="../barcode.asp?code=<%=(rsGetInvoice.Fields.Item("ID").Value)%>&height=30&width=1&mode=code39&text=0"><br><span style="font-size:18px;font-weight:bold"><%=(rsGetInvoice.Fields.Item("ID").Value)%></span>
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

<% if rsGetInvoice.Fields.Item("pay_method").Value = "Etsy" OR rsGetInvoice.Fields.Item("pay_method").Value = "Instagram" OR rsGetInvoice.Fields.Item("pay_method").Value = "Facebook" then %>
<center><h1><%= Ucase(rsGetInvoice("pay_method")) %> ORDER</h1>
<h3>Find thousands of more styles at Bodyartforms.com</h3>
</center>
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
objCmd.CommandText = "SELECT OrderDetailID, TBL_OrderSummary.qty, title, ProductDetail1, PreOrder_Desc, ProductDetails.price, notes, TBL_OrderSummary.item_price, DetailCode, ProductDetailID, anodization_id_ordered, anodization_fee, ID_Description, ID_SortOrder, ID_Number, type, autoclavable, BinNumber_Detail, Gauge, Length, location FROM TBL_OrderSummary INNER JOIN jewelry ON TBL_OrderSummary.ProductID = jewelry.ProductID INNER JOIN ProductDetails ON TBL_OrderSummary.DetailID = ProductDetails.ProductDetailID INNER JOIN sent_items ON TBL_OrderSummary.InvoiceID = sent_items.ID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE ID = ? ORDER BY ID_SortOrder ASC, BinNumber_Detail ASC, ProductDetailID ASC"
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
		<% if rsGetOrderDetails.Fields.Item("autoclavable").Value = 1 then %>
			<strong>
		<% end if %>
		<%=(rsGetOrderDetails.Fields.Item("title").Value)%>
				<% if rsGetOrderDetails.Fields.Item("autoclavable").Value = 1 then %>
			</strong>
		<% end if %>
		&nbsp;&nbsp;<%=(rsGetOrderDetails.Fields.Item("ProductDetail1").Value)%>&nbsp;<%=(rsGetOrderDetails.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetOrderDetails.Fields.Item("Length").Value)%>
		<% if rsGetOrderDetails("anodization_fee") > 0 then %>
		<span style="font-weight:bold;background-color:#d6d6d6;border-radius:5px;padding:3px">CUSTOM COLOR</span>
		<% end if %>
        <% if rsGetOrderDetails.Fields.Item("PreOrder_Desc").Value <> "" then %>
		 - 
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
	<% if var_gift_order <> "yes" then %>
    <td class="line-cost">
        <%= FormatCurrency((rsGetOrderDetails.Fields.Item("item_price").Value), -1, -2, -0, -2) %>
    </td>
    <td class="line-total">
        <%= FormatCurrency((rsGetOrderDetails.Fields.Item("item_price").Value)*(rsGetOrderDetails.Fields.Item("qty").Value), -1, -2, -0, -2) %>
		<% if rsGetOrderDetails("anodization_fee") > 0 then %>
		<br>
		 + <%= FormatCurrency(rsGetOrderDetails("qty") * rsGetOrderDetails("anodization_fee"),2) %> color add-on fee
		<% end if %>
    </td>
	<% end if %>
  </tr>
  <%
	LineItem = rsGetOrderDetails.Fields.Item("item_price").Value * rsGetOrderDetails.Fields.Item("qty").Value
	
	SumLineItem = SumLineItem + LineItem
	sum_anodization_fees = sum_anodization_fees + rsGetOrderDetails("qty") * rsGetOrderDetails("anodization_fee")
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
	<div class="qr-code" style="">
		<img src="/images/qr-codes/qr-return-policy.svg" style="width:100px;height:100;text-align: left;" />
	<div class="thank-you">
		Scan this QR code to view our return policy or if you have an order issue
	</div>
	<div class="our-promise">
		We want you to be 100% satisfied with your order! For order issues email us at help@bodyartforms.com or call (877) 223-5005.
	</div>
</td>
<td class="wrapper-totals">
	<table class="totals">
	<% if var_gift_order <> "yes" then %>
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
	arrTotals(2,2) = ""
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

	'arrTotals(0,6) = "Retail delivery fee" 
	'arrTotals(1,6) = "retail_delivery_fee"
	'retail_delivery_fee = rsGetInvoice.Fields.Item("retail_delivery_fee").Value
	'arrTotals(2,6) = ""

	For i = 0 to UBound(arrTotals, 2) 

		if rsGetInvoice.Fields.Item(arrTotals(1,i)).Value <> 0 then
	%>
	  <tr>
		<td class="line-cost"><%= arrTotals(0,i) %></td>
		<td class="line-total"><%= arrTotals(2,i) %>
		<%'uncoment below line to display delivery fee in the invoice%>
		<%'If (arrTotals(0,i) = "Tax" Or arrTotals(0,i) = "VAT") AND rsGetInvoice(arrTotals(1,6)) > 0 Then%>
			<%'= FormatCurrency(rsGetInvoice(arrTotals(1,i)) - rsGetInvoice(arrTotals(1,6)), -1, -2, -0, -2) %>
		<%'Else%>
			<%= FormatCurrency(rsGetInvoice(arrTotals(1,i)), -1, -2, -0, -2) %>
		<%'End If%>
		</td>
	  </tr>
	<% 
		end if ' if i > 2 or values not 0
	next ' loop through totals array


	InvoiceTotal = (SumLineItem + sum_anodization_fees - total_preferred_discount - total_coupon_discount - total_free_credits + rsGetInvoice.Fields.Item("shipping_rate").Value + total_sales_tax - total_store_credit - total_gift_cert)
	%>
	<% end if ' if gift order don't show pricing BUT DO show shipping method below 
	%>
	  <tr>
		<td class="line-cost"><%= (rsGetInvoice.Fields.Item("shipping_type").Value) %></td>
		<td class="line-total"><%= FormatCurrency((rsGetInvoice.Fields.Item("shipping_rate").Value), -1, -2, -0, -2)%></td>
	  </tr>
	  <% if var_gift_order <> "yes" then %>
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

<% else 'DOES have a return mailer 
%>

  <% end if ' DOES have a return mailer 
%>

</div>
</body>
</html>
<%
rsGetInvoice.Close()
%>