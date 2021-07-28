<% 
if free_count = 1 then
	var_free_amt = 30
elseif free_count = 2 then	
	var_free_amt = 50
elseif free_count = 3 then	
	var_free_amt = 75
elseif free_count = 4 then	
	var_free_amt = 100
elseif free_count = 5 then	
	var_free_amt = 150
end if
freeitem_found = ""
 
	' only show on CHECKOUT PAGE if a gift has been selected
	if gifts_checkout = "yes" and request.cookies("freegift" & free_count & "id") = "" then
		'do nothing
	else
	
	if (var_subtotal_after_discounts * 1 - var_totalvalue_certs_incart) < (var_free_amt * 1) then
		var_hide_gifts = "style=""visibility: hidden"""
	end if
	if var_other_items = 0 then
		var_hide_gifts = "style=""visibility: hidden"""
	end if
%>
<!--<div class="row freegift<%= free_count %>" <%= var_hide_gifts %>>-->
<div class="row freegift<%= free_count %>" id="gift<%= free_count %>" <%= var_hide_gifts %>>
<% if var_showgifts <> "no" then %>
<div class="dropdown w-100 my-1">
		<button class="btn w-100 btn-outline-secondary text-left dropdown-toggle" type="button" id="dropdownGift<%= free_count %>" data-flip="false" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">	  
				<span id="selected-gift<%= free_count %>">
					  <% Do While NOT rsGetFree.EOF 
					  if cStr(rsGetFree.Fields.Item("ProductDetailID").Value) = request.cookies("freegift" & free_count & "id") then
					  freeitem_found = "yes"
				  %>
					  <img class="ml-1 mr-2" style="width: 40px" src="https://s3.amazonaws.com/bodyartforms-products/<%= rsGetFree.Fields.Item("picture").Value %>"><span class="mr-3">Qty: <%=(rsGetFree.Fields.Item("Free_QTY").Value)%></span><%= Server.HTMLEncode(rsGetFree.Fields.Item("free_title").Value) %>
				  <% 
				  end if
				  rsGetFree.MoveNext()
				  Loop
				  rsGetFree.MoveFirst()
				  %>
				</span>
				<% if freeitem_found <> "yes" then %>
				<span id="gift<%= free_count %>-dropdown-text">Pick a free item</span>
				<% end if %>
			  </button>
			  <div class="dropdown-menu w-100 modal-scroll-long" aria-labelledby="dropdownGift<%= free_count %>">
				  <div class="dropdown-item btn-group-vertical btn-group-toggle m-0 p-0 " data-toggle="buttons">
					  <label class="btn btn-light d-block text-left">
						  <input type="radio" name="freegift<%= free_count %>" id="freegift0" value="" data-friendly="No free item" data-img-name="blank.gif"><img class="ml-1 mr-2" style="width: 40px" src="https://s3.amazonaws.com/bodyartforms-products/blank.gif">I don't need a free item
					  </label>
<% Do While Not rsGetFree.EOF
if rsGetFree.Fields.Item("free").Value <= var_free_amt then 
	display_option = "yes"
	
	if cStr(rsGetFree.Fields.Item("ProductDetailID").Value) = request.cookies("freegift" & free_count & "id") then
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
	end if ' hide SAVE FOR LATER credit if user is not logged in
	
	if display_option = "yes" then

 %>
 <label class="btn btn-light d-block text-left">
		<input type="radio" name="freegift<%= free_count %>" id="freegift<%= free_count %>" value="<%= rsGetFree.Fields.Item("ProductDetailID").Value %>" data-friendly="<span class='mr-3'>Qty: <%=(rsGetFree.Fields.Item("Free_QTY").Value)%></span><%= Server.HTMLEncode(rsGetFree.Fields.Item("free_title").Value) %>" data-img-name="<%= rsGetFree.Fields.Item("picture").Value %>"><img class="ml-1 mr-2" style="width: 40px" src="https://s3.amazonaws.com/bodyartforms-products/<%= rsGetFree.Fields.Item("picture").Value %>">
		<span class="mr-3">Qty: <%=(rsGetFree.Fields.Item("Free_QTY").Value)%></span><%= Server.HTMLEncode(rsGetFree.Fields.Item("free_title").Value) %>
   </label> 
  <%
	end if ' display option = "yes"
end if ' only get free items valued at free amount or below  
  rsGetFree.MoveNext()
Loop
	rsGetFree.MoveFirst()
	
%>  
</div><!-- button group -->
</div><!-- drop down menu -->
</div><!-- drop down -->
<!--</select> -->
<% else ' don't allow user to change selection on checkout page 

	rsGetFree.MoveFirst()
 end if ' don't allow user to change selection on checkout page 
 %>
</div><!-- end free item row -->
 <%
			
'set variables to use on final page of checkout processing
session("credit_now") = credit_now
session("credit_later") = credit_later
%>
<!--</div>--><!-- end row -->
<%
	' only show on CHECKOUT PAGE if a gift has been selected
	end if
%>