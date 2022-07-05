<%

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM jewelry WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,request("productid")))
	Set rsProduct = objCmd.Execute()

	if instr(rsProduct.Fields.Item("jewelry").Value, "captive") OR instr(rsProduct.Fields.Item("jewelry").Value, "hoop") OR instr(rsProduct.Fields.Item("jewelry").Value, "clicker") OR instr(rsProduct.Fields.Item("jewelry").Value, "pinchers") then
		var_length = "Diameter"
	else
		var_length = "Length"
	end if
	
if request("gauge") <> "" then
	filter_gauge = " and gauge = ? "
end if

'===== only used from cart page to not allow items to show up to select from that do not have enough qty in stock ======
if request("cart_qty") <> "" then
	filter_cart_qty = " and ? <= ProductDetails.qty "
	var_cart_modal = "width:100%"
else '==== standard product details page====
	filter_cart_qty = " and qty > 0 "
end if


set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT (SELECT ProductDetails.qty - COALESCE(SUM(cart_qty), 0) FROM tbl_carts WHERE ProductDetails.ProductDetailID = tbl_carts.cart_detailID AND cart_dateAdded > DATEADD(mi, -60, GETDATE())) as dynamic_qty," & _ 
		"L.length_mm as 'length_conversion', *, " & _
		"ISNULL(Gauge, '') + ' ' + ISNULL(Length, '') + ' ' + ISNULL(ProductDetail1, '') as OptionTitle," & _ 
		"ISNULL(Gauge, ''), ISNULL(Length, ''), " & _ 
		"ISNULL(ProductDetail1, '') " & _ 
		"FROM ProductDetails INNER JOIN TBL_GaugeOrder as G ON ISNULL(ProductDetails.Gauge,'') = ISNULL(G.GaugeShow,'') " & _ 
		"LEFT JOIN tbl_lengths as L ON ISNULL(ProductDetails.Length,'') = ISNULL(L.length_inches,'') " & _ 
		"WHERE ProductID = ? AND active = 1 " & filter_gauge & " " & filter_cart_qty & " ORDER BY G.GaugeOrder ASC, item_order ASC, Price ASC"

objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,request("productid")))
if request("gauge") <> "" then
	objCmd.Parameters.Append(objCmd.CreateParameter("gauge",200,1,20,request("gauge")))
end if
if request("cart_qty") <> "" then
	objCmd.Parameters.Append(objCmd.CreateParameter("cart_qty",3,1,20, request("cart_qty")))
end if

set rsGetItems = Server.CreateObject("ADODB.Recordset")
rsGetItems.CursorLocation = 3 'adUseClient
rsGetItems.Open objCmd
var_totalitems = rsGetItems.RecordCount

if not rsGetItems.eof then
%>
<div class="dropdown my-2" id="add-cart-menu" style="<%= var_cart_modal %>">
<button class="btn btn-light rounded-0 bg-white text-left dropdown-toggle font-weight-bold  py-2" style="border:1px solid #ced4da" type="button" id="dropdownAddCart" data-flip="false" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false" style="<%= var_cart_modal %>">
	<%If var_totalitems = 1 Then%>
		<span id="selected-item"></span>
	<%Else%>
		<span id="selected-item">Select item:</span>
	<%End If%>
</button>
<div class="dropdown-menu modal-scroll-long rounded-0" style="border:2px solid #ced4da" aria-labelledby="dropdownAddCart">
		<div id="msg-filtered-dropdown"></div>
<div class="dropdown-item bg-white btn-group-vertical btn-group-toggle m-0 p-0 " data-toggle="buttons">

			<%
			optgauge = ""
			i_count = 0
			While NOT rsGetItems.EOF

			if rsGetItems.Fields.Item("qty").Value > "0" then 
			
			' Set drop down variables
			option_sale_price = 0
			option_pair = " (each) "
			option_retail_price = FormatNumber(rsGetItems.Fields.Item("price").Value  * exchange_rate , -1, -2, -0, -2)
			option_description = rsGetItems.Fields.Item("OptionTitle").Value

			if (rsProduct.Fields.Item("SaleDiscount").Value > 0 AND rsProduct.Fields.Item("secret_sale").Value = 0) OR  (rsProduct.Fields.Item("secret_sale").Value = 1 AND session("secret_sale") = "yes") then
				option_sale_price = FormatNumber((rsGetItems.Fields.Item("price").Value * exchange_rate/100) * (100 - rsProduct.Fields.Item("SaleDiscount").Value), -1, -2, -2, -2)
				
				option_actual_price = formatnumber(option_sale_price,2)
			else
				option_actual_price = formatnumber(option_retail_price,2)
			end if
		
			if sale_retail_price = "" then
				sale_retail_price = option_retail_price
				sale_savings = FormatNumber(option_retail_price - option_sale_price, -1, -2, -2, -2)
			end if

			if rsProduct.Fields.Item("pair").Value = "yes" then
			option_pair = " (pair) "
			end if
			
			var_qty_in_stock = ""
			if rsProduct.Fields.Item("customorder").Value <> "yes" and rsProduct.Fields.Item("type").Value <> "One time buy" and rsGetItems.Fields.Item("dynamic_qty").Value <= 4 then
				var_qty_in_stock = "&nbsp;&nbsp;&nbsp;[" & rsGetItems.Fields.Item("dynamic_qty").Value & " left]"
			end if
						
			setgroup = 0
			' add optgroup tag for larger listings
			if var_totalitems > 20 then
				if rsGetItems.Fields.Item("gauge").Value <> optgauge then
				setgroup = 1
			%>
			<div class="font-weight-bold bg-dark text-light w-100 pl-2 py-1 mt-2">
				<%= Server.HTMLEncode(rsGetItems.Fields.Item("gauge").Value) %> OPTIONS
			</div>
			<%	
				end if 
			end if
			%>
			<label class="btn rounded-0 py-3 py-lg-2 border-bottom text-left btn-select-menu option_img_<%=(rsGetItems.Fields.Item("img_id").Value)%>">
				<input class="add-cart" type="radio" name="add-cart" value="<%=(rsGetItems.Fields.Item("ProductDetailID").Value)%>" data-qty="<%=rsGetItems.Fields.Item("dynamic_qty").Value%>" data-img_id="<%=(rsGetItems.Fields.Item("img_id").Value)%>" data-sale-price="<%= option_sale_price %>" data-retail-price="<%= option_retail_price %>" data-actual-price="<%= option_actual_price %>" data-symbol="<%= exchange_symbol %>" data-title="<%= replace(rsGetItems.Fields.Item("ProductDetail1").Value, """", "") %>" dropdown-title="<%= exchange_symbol %><%= option_actual_price %>
				&nbsp;&nbsp;&nbsp;&nbsp;<%= server.htmlencode(rsGetItems.Fields.Item("OptionTitle").Value) %>" data-variant="<%= trim(server.htmlencode(rsGetItems.Fields.Item("OptionTitle").Value)) %>" required   <%if var_totalitems = 1 Then Response.Write "checked"%>><%= exchange_symbol %><%= option_actual_price %>
				&nbsp;&nbsp;&nbsp;&nbsp;
				<% if rsGetItems.Fields.Item("Gauge").Value <> "" then %>
					<%= rsGetItems.Fields.Item("Gauge").Value %>&nbsp;
				<% if request.Cookies("showmm") = "yes" and rsGetItems.Fields.Item("display_conversion").Value = 1 then %>(<%= rsGetItems.Fields.Item("conversion").Value %>) 
				<% end if %>
				<% end if %>
				<% if rsGetItems.Fields.Item("length").Value <> "" then %>
				&#9679; <%= var_length %>: <%= rsGetItems.Fields.Item("Length").Value %> 
				<% if request.Cookies("showmm") = "yes" and rsGetItems.Fields.Item("length_conversion").Value <> "" then %> (<%= rsGetItems.Fields.Item("length_conversion").Value %>) 
				<% end if %>
				&#9679;
				<% end if %>

				
				&nbsp;<%= rsGetItems.Fields.Item("ProductDetail1").Value %>
				<%= var_qty_in_stock %>
				</label> 
			<%
			optgauge = rsGetItems.Fields.Item("gauge").Value
			
			end if ' qty > 0
			
			i_count = i_count + 1
			rsGetItems.MoveNext()
			Wend 
			rsGetItems.Requery() 
			%>

		</div><!-- button group -->
		</div><!-- drop down menu -->
		</div><!-- drop down -->
<!-- for cart page update -->
<input type="hidden" name="cartid" value="<%= request("cartid") %>">
<%
end if 	' not rsGetItems.eof 
%>
<script>
	// If there's only one option in the list copy the default checked value and display it to user
	var ele = document.getElementsByName('add-cart');
		
	for(i = 0; i < ele.length; i++) {
		if(ele[i].checked)
		document.getElementById("selected-item").innerHTML
				= ele[i].getAttribute('dropdown-title');			
	}
</script>