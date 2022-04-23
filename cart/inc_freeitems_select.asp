<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
productid = Request("productid")
tier = Request("tier")

Select Case tier
	Case 1
		var_free_amt = 30
	Case 2
		var_free_amt = 30		
	Case 3
		var_free_amt = 30
	Case 4
		var_free_amt = 50
	Case 5
		var_free_amt = 75
	Case 6
		var_free_amt = 100
	Case 7
		var_free_amt = 150	
End Select

If tier = 1 Then
	' ------- Get O_RINGS items
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT ProductDetails.free, 4 As Free_QTY, jewelry.picture, ProductDetails.ProductDetailID, ProductDetails.ProductDetail1, FlatProducts.min_gauge, FlatProducts.max_gauge, jewelry.title, jewelry.picture,jewelry.ProductID, jewelry.picture, CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END, ISNULL(ProductDetails.gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' + ISNULL(ProductDetails.ProductDetail1,'') + ' ' + ISNULL(jewelry.title,'') AS 'free_title' " & _
						"FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID " & _ 
						"INNER JOIN FlatProducts ON FlatProducts.ProductID = jewelry.ProductID " & _ 
						"INNER JOIN TBL_GaugeOrder Gauge ON ISNULL(ProductDetails.Gauge,'') = ISNULL(Gauge.GaugeShow,'') " & _ 
						"WHERE (jewelry.ProductID = " & productid & ") AND ProductDetails.qty > 0 " & _
						"ORDER BY GaugeOrder ASC, item_order ASC, Price ASC"	
					
	Set rsGetFree = objCmd.Execute()
	' ------- End getting O_RINGS items
End If

If tier = 2 Then
	'' ------- Get STICKER items
	'set objCmd = Server.CreateObject("ADODB.command")
	'objCmd.ActiveConnection = DataConn
	'objCmd.CommandText = "SELECT ProductDetails.free, ProductDetails.Free_QTY, jewelry.picture, ProductDetails.ProductDetailID, ProductDetails.ProductDetail1, FlatProducts.min_gauge, FlatProducts.max_gauge, jewelry.title, jewelry.picture,jewelry.ProductID, jewelry.picture, CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END, ISNULL(ProductDetails.gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' + ISNULL(ProductDetails.ProductDetail1,'') + ' ' + ISNULL(jewelry.title,'') AS 'free_title' " & _
	'					"FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID " & _ 
	'					"INNER JOIN FlatProducts ON FlatProducts.ProductID = jewelry.ProductID " & _ 
	'					"INNER JOIN TBL_GaugeOrder Gauge ON ISNULL(ProductDetails.Gauge,'') = ISNULL(Gauge.GaugeShow,'') " & _ 
	'					"WHERE (jewelry.ProductID = " & productid & ") AND (ProductDetails.qty > 0) AND (ProductDetails.free <> 0) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.active = 1) " & _
	'					"ORDER BY GaugeOrder ASC, item_order ASC, Price ASC"
	'Set rsGetFree = objCmd.Execute()
	'' ------- End getting STICKER items
End If

If tier >= 3 Then
	' ------- Get FREE items
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT ProductDetails.free, ProductDetails.Free_QTY, jewelry.picture, ProductDetails.ProductDetailID, ProductDetails.ProductDetail1, FlatProducts.min_gauge, FlatProducts.max_gauge, jewelry.title, jewelry.picture,jewelry.ProductID, jewelry.picture, CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END, ISNULL(ProductDetails.gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' + ISNULL(ProductDetails.ProductDetail1,'') + ' ' + ISNULL(jewelry.title,'') AS 'free_title' " & _
						"FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID " & _ 
						"INNER JOIN FlatProducts ON FlatProducts.ProductID = jewelry.ProductID " & _ 
						"INNER JOIN TBL_GaugeOrder Gauge ON ISNULL(ProductDetails.Gauge,'') = ISNULL(Gauge.GaugeShow,'') " & _ 
						"WHERE (jewelry.ProductID = " & productid & ") AND (ProductDetails.qty > 0) AND (ProductDetails.free <> 0) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.active = 1) " & _
						"ORDER BY GaugeOrder ASC, item_order ASC, Price ASC"
	Set rsGetFree = objCmd.Execute()
	' ------- End getting free items
End If

%>
<div class="row" style="">
	<div class="dropdown w-100 my-1" style="">
		<button class="btn w-100 text-left dropdown-toggle" type="button" id="dropdownGift<%=tier%>" data-flip="false" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false" style="-webkit-box-shadow: 0 0 0 0.3rem rgb(130 138 145 / 50%); box-shadow: 0 0 0 0.3rem rgb(130 138 145 / 50%); background-color: #3d454c; color: white;">	  
			<span id="selected-gift<%= tier %>">
				<% 
				rsGetFree.MoveFirst()
				Do While NOT rsGetFree.EOF 
					if cStr(rsGetFree.Fields.Item("ProductDetailID").Value) = Cstr(request.cookies("freegift" & tier & "id")) then
						freeitem_found = "yes"
						Response.Write Server.HTMLEncode(rsGetFree.Fields.Item("free_title").Value)
					end if
					rsGetFree.MoveNext()
				Loop
				rsGetFree.MoveFirst()
				%>
			</span>
			<% if freeitem_found <> "yes" then %>
				<span id="gift<%= tier %>-dropdown-text">Select item</span>
			<% end if %>
		</button>
		<div class="dropdown-menu w-100 modal-scroll-long" aria-labelledby="dropdownGift<%=tier%>" style="z-index:2000">
			<div class="dropdown-item btn-group-vertical btn-group-toggle m-0 p-0 " data-toggle="buttons">
				<label class="btn btn-light d-block text-left">
				  <input type="radio" class="freegift" data-tier="<%= tier %>" name="freegift<%= tier %>" id="freegift0" value="" data-slide-index="<%=Request("slideindex")%>" data-friendly="no free item">I don't need a free item
				</label>
				<% Do While Not rsGetFree.EOF

					display_option = "yes"

					if cStr(rsGetFree.Fields.Item("ProductDetailID").Value) = Cstr(request.cookies("freegift" & tier & "id")) then
						var_selected = "selected"
					else
						var_selected = ""
					end if

					' only show credit options applicable to that free value range and not all the older ones
					if rsGetFree.Fields.Item("ProductID").Value = 2890 and rsGetFree.Fields.Item("free").Value <> var_free_amt then
						display_option = "no"
					end if

					' hide SAVE FOR LATER credit if user is not logged in
					if CustID_Cookie = 0 and Instr(1, rsGetFree.Fields.Item("ProductDetail1"), "LATER") > 0 then
						display_option = "no"
					end if 

					if display_option = "yes" then
						%>
						<label class="btn btn-light d-block text-left">
							<input type="radio" class="freegift" data-tier="<%= tier %>" name="freegift<%= tier %>" id="freegift<%= tier %>" value="<%= rsGetFree.Fields.Item("ProductDetailID").Value %>" data-slide-index="<%=Request("slideindex")%>" data-friendly="<%= Server.HTMLEncode(rsGetFree.Fields.Item("free_title").Value) %>" data-img-name="<%= rsGetFree.Fields.Item("picture").Value %>">
							<span class="mr-3">Qty: <%=(rsGetFree.Fields.Item("Free_QTY").Value)%></span><%= Server.HTMLEncode(rsGetFree.Fields.Item("free_title").Value) %>
						</label> 
						<%
					end if 
					rsGetFree.MoveNext()
				Loop
				%>  
			</div>
		</div>
	</div>
</div>
<%
Set rsGetFree = Nothing		
'set variables to use on final page of checkout processing
session("credit_now") = credit_now
session("credit_later") = credit_later

%>